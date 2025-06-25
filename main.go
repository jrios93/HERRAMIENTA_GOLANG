package main

import (
	"bytes"
	"fmt"
	"io/ioutil"
	"log"
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"time"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/canvas"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/storage"
	"fyne.io/fyne/v2/widget"
	"github.com/go-vgo/robotgo"
	"github.com/jung-kurt/gofpdf"
	hook "github.com/robotn/gohook"
	"github.com/skip2/go-qrcode"
)

var cancel = make(chan struct{})

const (
	saveFile         = "bloc_notas.txt"
	autoSaveInterval = 5 * time.Second

	// Rutas para los logos
	logosDir     = "logos"
	zettacomLogo = "logos/zettacom.png"
	comsitecLogo = "logos/comsitec.png"

	// Fuentes
	fontsDir = "fonts"
)

// Datos predefinidos de empresas
var empresasData = map[string]struct {
	Nombre    string
	Direccion string
	Telefono  string
	NeedQR    bool
	Color     struct{ R, G, B int }
}{
	"ZETTACOM": {
		Nombre:    "ZETTACOM S.A.C",
		Direccion: "Av. Giraldez 242, Huancayo, Junín",
		Telefono:  "+51 964 789 123",
		NeedQR:    false,
		Color:     struct{ R, G, B int }{0, 51, 102}, // Azul marino
	},
	"COMSITEC": {
		Nombre:    "COMSITEC S.A.C",
		Direccion: "Av. Giraldez 242, Huancayo, Junín",
		Telefono:  "+51 964 789 456",
		NeedQR:    true,
		Color:     struct{ R, G, B int }{180, 20, 40}, // Rojo corporativo
	},
}

// Tamaños de papel en mm
var paperSizes = map[string]struct {
	Width  float64
	Height float64
}{
	"A4":    {Width: 210, Height: 297},
	"A5":    {Width: 148, Height: 210},
	"Carta": {Width: 216, Height: 279},
}

type Item struct {
	Codigo string
	Nombre string
	Firma  string
}

type NotePad struct {
	multiLine    *widget.Entry
	lastContent  string
	lastSaveTime time.Time
	statusLabel  *widget.Label
	lastUserEdit time.Time
}

type RotuloData struct {
	Empresa               string
	RemitenteNombre       string
	RemitenteDireccion    string
	RemitenteTelefono     string
	DestinatarioNombre    string
	DestinatarioDireccion string
	DestinatarioTelefono  string
	Peso                  string
	Observaciones         string
	NumeroGuia            string
	TamanoHoja            string
	Orientacion           string
	FechaEnvio            time.Time
}

type RotuloGenerator struct {
	data         *RotuloData
	preview      *widget.RichText
	empresaCheck *widget.RadioGroup
	inputs       map[string]*widget.Entry
	tamanoHoja   *widget.Select
	orientacion  *widget.RadioGroup
	logoPreview  *canvas.Image
	pdfPreview   *widget.Label
	window       fyne.Window
	pdfCounter   int
}

func main() {
	a := app.New()
	w := a.NewWindow("Mi herramienta de trabajo")
	w.Resize(fyne.NewSize(1200, 700))

	// Crear directorios necesarios
	createRequiredDirs()

	// Tab 1: Autocopiador
	autocopiadorTab := createAutocopiadorTab(w)

	// Tab 2: Personal
	notepad := &NotePad{}
	personalTab := notepad.createPersonalTab(w)

	// Tab 3: Rótulo Profesional
	rotuloGenerator := &RotuloGenerator{
		data: &RotuloData{
			TamanoHoja:  "A4",
			Orientacion: "Vertical",
			FechaEnvio:  time.Now(),
		},
		inputs:     make(map[string]*widget.Entry),
		window:     w,
		pdfCounter: 1,
	}
	rotuloTab := rotuloGenerator.createRotuloTab(w)

	tabs := container.NewAppTabs(
		container.NewTabItem("🤖 Autocopiador", autocopiadorTab),
		container.NewTabItem("📝 Personal", personalTab),
		container.NewTabItem("🏷️ Rótulo Profesional", rotuloTab),
	)

	w.SetContent(tabs)
	w.Show()

	go globalEscapeListener(nil)
	a.Run()
}

func createRequiredDirs() {
	// Crear directorio para logos si no existe
	if _, err := os.Stat(logosDir); os.IsNotExist(err) {
		os.Mkdir(logosDir, 0755)
		fmt.Printf("Directorio para logos creado: %s\n", logosDir)
		fmt.Printf("Por favor, coloca tus archivos de logo como:\n- %s\n- %s\n", zettacomLogo, comsitecLogo)
	}

	// Crear directorio para fuentes si no existe
	if _, err := os.Stat(fontsDir); os.IsNotExist(err) {
		os.Mkdir(fontsDir, 0755)
		fmt.Printf("Directorio para fuentes creado: %s\n", fontsDir)
	}
}

func createAutocopiadorTab(window fyne.Window) *fyne.Container {
	// Input de series
	seriesInput := widget.NewMultiLineEntry()
	seriesInput.SetPlaceHolder("Ejemplo: 12345 67890 11111 22222\n(Separa las series con espacios)")

	seriesScroll := container.NewScroll(seriesInput)
	seriesScroll.SetMinSize(fyne.NewSize(480, 180))

	dateInput := widget.NewEntry()
	dateInput.SetPlaceHolder("Formato: 15052025 (DDMMAAAA)")

	// Labels de estado
	statusLabel := widget.NewLabel("Estado: Esperando acción...")
	statusLabel.Importance = widget.MediumImportance

	copiedCounter := widget.NewLabel("Copiadas: 0 / 0")
	copiedCounter.Importance = widget.LowImportance

	// Botones
	startButton := widget.NewButton("▶️ Iniciar Autocopiado", func() {
		rawSeries := seriesInput.Text
		date := dateInput.Text

		if strings.TrimSpace(rawSeries) == "" {
			dialog.ShowError(fmt.Errorf("debes ingresar al menos una serie"), window)
			return
		}
		if strings.TrimSpace(date) == "" {
			dialog.ShowError(fmt.Errorf("debes ingresar una fecha"), window)
			return
		}

		delayMs := 90
		countdownSec := 5

		statusLabel.SetText(fmt.Sprintf("Iniciando en %d segundos...", countdownSec))
		copiedCounter.SetText("Copiadas: 0 / 0")

		cancel = make(chan struct{})

		go autocopiar(rawSeries, date, time.Duration(delayMs)*time.Millisecond, countdownSec, statusLabel, copiedCounter)
	})
	startButton.Importance = widget.HighImportance

	cancelButton := widget.NewButton("⏹️ Cancelar", func() {
		select {
		case <-cancel:
		default:
			close(cancel)
			statusLabel.SetText("Estado: Cancelado manualmente.")
		}
	})
	cancelButton.Importance = widget.MediumImportance

	// Información de ayuda
	helpText := widget.NewRichTextFromMarkdown(`
**Instrucciones:**
1. Ingresa las series separadas por espacios
2. Ingresa la fecha en formato DDMMAAAA
3. Presiona "Iniciar Autocopiado"
4. Puedes cancelar con el botón o presionando ESC

**Nota:** El proceso comenzará después de una cuenta regresiva de 5 segundos.
`)
	helpText.Wrapping = fyne.TextWrapWord

	helpScroll := container.NewScroll(helpText)
	helpScroll.SetMinSize(fyne.NewSize(350, 120))

	// Cards
	inputCard := widget.NewCard("📋 Datos de Entrada", "",
		container.NewVBox(
			widget.NewLabel("Series:"),
			seriesScroll,
			widget.NewLabel("Fecha:"),
			dateInput,
		),
	)

	controlCard := widget.NewCard("🎮 Controles", "",
		container.NewVBox(
			container.NewHBox(startButton, cancelButton),
			widget.NewSeparator(),
			statusLabel,
			copiedCounter,
		),
	)

	helpCard := widget.NewCard("ℹ️ Ayuda", "", helpScroll)

	return container.NewVBox(
		widget.NewLabel("Autocopiador de Series"),
		container.NewHBox(
			container.NewVBox(inputCard, controlCard),
			helpCard,
		),
	)
}

func (r *RotuloGenerator) createRotuloTab(window fyne.Window) *fyne.Container {
	// Inicializar vista previa
	r.preview = widget.NewRichText()
	r.preview.Wrapping = fyne.TextWrapWord

	// Selección de empresa
	r.empresaCheck = widget.NewRadioGroup([]string{"ZETTACOM", "COMSITEC"}, func(selected string) {
		r.data.Empresa = selected

		// Autocompletar datos
		if empresaData, ok := empresasData[selected]; ok {
			r.inputs["remitenteNombre"].SetText(empresaData.Nombre)
			r.inputs["remitenteDireccion"].SetText(empresaData.Direccion)
			r.inputs["remitenteTelefono"].SetText(empresaData.Telefono)
		}

		r.updateLogoPreview(selected)
		r.updatePreview()
	})
	r.empresaCheck.Horizontal = true

	// Logo preview
	r.logoPreview = &canvas.Image{}
	r.logoPreview.Resize(fyne.NewSize(150, 80))
	r.logoPreview.FillMode = canvas.ImageFillContain

	// Configuración
	r.tamanoHoja = widget.NewSelect(
		[]string{"A4", "A5", "Carta"},
		func(selected string) {
			r.data.TamanoHoja = selected
			r.updatePreview()
		},
	)
	r.tamanoHoja.SetSelected("A4")

	r.orientacion = widget.NewRadioGroup(
		[]string{"Vertical", "Horizontal"},
		func(selected string) {
			r.data.Orientacion = selected
			r.updatePreview()
		},
	)
	r.orientacion.Horizontal = true
	r.orientacion.SetSelected("Vertical")

	// Crear inputs
	r.createInputs()

	// Botones de acción
	generateButton := widget.NewButton("📄 Generar Rótulo PDF", func() {
		r.generateProfessionalPDF(window)
	})
	generateButton.Importance = widget.HighImportance

	printButton := widget.NewButton("🖨️ Imprimir", func() {
		r.printRotulo(window)
	})
	printButton.Importance = widget.MediumImportance

	clearButton := widget.NewButton("🗑️ Limpiar", func() {
		r.clearFields()
	})

	autoFillButton := widget.NewButton("🔄 Datos de Prueba", func() {
		r.fillTestData()
	})

	// Vista previa
	previewScroll := container.NewScroll(r.preview)
	previewScroll.SetMinSize(fyne.NewSize(400, 500))

	// Layout del formulario
	formCard := r.createFormLayout()

	// Card de vista previa
	previewCard := widget.NewCard("👁️ Vista Previa del Rótulo", "", previewScroll)

	// Card de controles
	controlCard := widget.NewCard("🎮 Acciones", "",
		container.NewVBox(
			container.NewGridWithColumns(2, generateButton, printButton),
			container.NewGridWithColumns(2, autoFillButton, clearButton),
			widget.NewSeparator(),
			widget.NewLabel("✨ Rótulo profesional con logo y QR"),
			widget.NewLabel("📦 Diseño adaptado al tamaño seleccionado"),
			widget.NewLabel("🔍 Soporte para caracteres especiales"),
		),
	)

	// Establecer valores por defecto
	r.empresaCheck.SetSelected("ZETTACOM")
	r.data.Empresa = "ZETTACOM"
	r.updateLogoPreview("ZETTACOM")
	r.updatePreview()

	// Layout principal
	formScroll := container.NewScroll(formCard)
	formScroll.SetMinSize(fyne.NewSize(600, 500))

	return container.NewVBox(
		container.NewHBox(
			formScroll,
			container.NewVBox(previewCard, controlCard),
		),
	)
}

func (r *RotuloGenerator) createInputs() {
	r.inputs["remitenteNombre"] = widget.NewEntry()
	r.inputs["remitenteNombre"].SetPlaceHolder("Nombre completo del remitente")
	r.inputs["remitenteNombre"].OnChanged = func(text string) {
		r.data.RemitenteNombre = text
		r.updatePreview()
	}

	r.inputs["remitenteDireccion"] = widget.NewMultiLineEntry()
	r.inputs["remitenteDireccion"].SetPlaceHolder("Dirección completa del remitente")
	r.inputs["remitenteDireccion"].Resize(fyne.NewSize(300, 60))
	r.inputs["remitenteDireccion"].OnChanged = func(text string) {
		r.data.RemitenteDireccion = text
		r.updatePreview()
	}

	r.inputs["remitenteTelefono"] = widget.NewEntry()
	r.inputs["remitenteTelefono"].SetPlaceHolder("Teléfono del remitente")
	r.inputs["remitenteTelefono"].OnChanged = func(text string) {
		r.data.RemitenteTelefono = text
		r.updatePreview()
	}

	r.inputs["destinatarioNombre"] = widget.NewEntry()
	r.inputs["destinatarioNombre"].SetPlaceHolder("Nombre completo del destinatario")
	r.inputs["destinatarioNombre"].OnChanged = func(text string) {
		r.data.DestinatarioNombre = text
		r.updatePreview()
	}

	r.inputs["destinatarioDireccion"] = widget.NewMultiLineEntry()
	r.inputs["destinatarioDireccion"].SetPlaceHolder("Dirección completa del destinatario")
	r.inputs["destinatarioDireccion"].Resize(fyne.NewSize(300, 60))
	r.inputs["destinatarioDireccion"].OnChanged = func(text string) {
		r.data.DestinatarioDireccion = text
		r.updatePreview()
	}

	r.inputs["destinatarioTelefono"] = widget.NewEntry()
	r.inputs["destinatarioTelefono"].SetPlaceHolder("Teléfono del destinatario")
	r.inputs["destinatarioTelefono"].OnChanged = func(text string) {
		r.data.DestinatarioTelefono = text
		r.updatePreview()
	}

	r.inputs["peso"] = widget.NewEntry()
	r.inputs["peso"].SetPlaceHolder("Peso del paquete (opcional)")
	r.inputs["peso"].OnChanged = func(text string) {
		r.data.Peso = text
		r.updatePreview()
	}

	r.inputs["numeroGuia"] = widget.NewEntry()
	r.inputs["numeroGuia"].SetPlaceHolder("Número de guía (se genera automático)")
	r.inputs["numeroGuia"].OnChanged = func(text string) {
		r.data.NumeroGuia = text
		r.updatePreview()
	}

	r.inputs["observaciones"] = widget.NewMultiLineEntry()
	r.inputs["observaciones"].SetPlaceHolder("Observaciones especiales")
	r.inputs["observaciones"].Resize(fyne.NewSize(300, 60))
	r.inputs["observaciones"].OnChanged = func(text string) {
		r.data.Observaciones = text
		r.updatePreview()
	}
}

func (r *RotuloGenerator) createFormLayout() *widget.Card {
	// Empresa y logo
	empresaForm := container.NewVBox(
		widget.NewLabel("EMPRESA:"),
		r.empresaCheck,
		container.NewCenter(r.logoPreview),
	)

	// Remitente
	remitenteForm := container.NewVBox(
		widget.NewLabel("REMITENTE:"),
		widget.NewLabel("Nombre:"),
		r.inputs["remitenteNombre"],
		widget.NewLabel("Dirección:"),
		r.inputs["remitenteDireccion"],
		widget.NewLabel("Teléfono:"),
		r.inputs["remitenteTelefono"],
	)

	// Destinatario
	destinatarioForm := container.NewVBox(
		widget.NewLabel("DESTINATARIO:"),
		widget.NewLabel("Nombre:"),
		r.inputs["destinatarioNombre"],
		widget.NewLabel("Dirección:"),
		r.inputs["destinatarioDireccion"],
		widget.NewLabel("Teléfono:"),
		r.inputs["destinatarioTelefono"],
	)

	// Detalles
	detallesForm := container.NewVBox(
		widget.NewLabel("DETALLES DEL ENVÍO:"),
		container.NewGridWithColumns(2,
			container.NewVBox(
				widget.NewLabel("Peso (opcional):"),
				r.inputs["peso"],
			),
			container.NewVBox(
				widget.NewLabel("Número de Guía:"),
				r.inputs["numeroGuia"],
			),
		),
		widget.NewLabel("Observaciones:"),
		r.inputs["observaciones"],
	)

	// Configuración
	configForm := container.NewVBox(
		widget.NewLabel("CONFIGURACIÓN:"),
		container.NewGridWithColumns(2,
			container.NewVBox(
				widget.NewLabel("Tamaño:"),
				r.tamanoHoja,
			),
			container.NewVBox(
				widget.NewLabel("Orientación:"),
				r.orientacion,
			),
		),
		widget.NewLabel("💡 El diseño se adaptará automáticamente"),
		widget.NewLabel("📄 Todo el contenido en una sola página"),
	)

	return widget.NewCard("📋 Datos del Envío", "",
		container.NewVBox(
			empresaForm,
			widget.NewSeparator(),
			container.NewGridWithColumns(2, remitenteForm, destinatarioForm),
			widget.NewSeparator(),
			detallesForm,
			widget.NewSeparator(),
			configForm,
		),
	)
}

func (r *RotuloGenerator) generateProfessionalPDF(window fyne.Window) {
	if r.data.RemitenteNombre == "" || r.data.DestinatarioNombre == "" {
		dialog.ShowError(fmt.Errorf("debes completar al menos el nombre del remitente y destinatario"), window)
		return
	}

	// Generar número de guía si está vacío
	if r.data.NumeroGuia == "" {
		r.data.NumeroGuia = fmt.Sprintf("%s%d", r.data.Empresa[:3], time.Now().Unix()%1000000)
	}

	timestamp := time.Now().Format("20060102_150405")
	defaultName := fmt.Sprintf("rotulo_%s_%s_%s.pdf", r.data.Empresa, r.data.NumeroGuia, timestamp)

	saveDialog := dialog.NewFileSave(
		func(writer fyne.URIWriteCloser, err error) {
			if err != nil {
				dialog.ShowError(err, window)
				return
			}
			if writer == nil {
				return
			}
			defer writer.Close()

			// Generar PDF profesional
			pdfData, err := r.createProfessionalPDF()
			if err != nil {
				dialog.ShowError(fmt.Errorf("error generando PDF: %v", err), window)
				return
			}

			_, writeErr := writer.Write(pdfData)
			if writeErr != nil {
				dialog.ShowError(writeErr, window)
				return
			}

			r.pdfCounter++
			filePath := writer.URI().Path()

			dialog.ShowInformation("✅ Rótulo Generado",
				fmt.Sprintf("Rótulo profesional generado exitosamente:\n\n"+
					"📄 Archivo: %s\n"+
					"🏢 Empresa: %s\n"+
					"📦 Guía: %s\n"+
					"📏 Tamaño: %s - %s\n"+
					"👤 Remitente: %s\n"+
					"📍 Destinatario: %s\n\n"+
					"✨ Incluye:\n"+
					"• Logo corporativo\n"+
					"• Código de barras\n"+
					"• Diseño adaptado al tamaño\n"+
					"• Soporte para caracteres especiales\n"+
					"• Todo en una sola página",
					filepath.Base(filePath),
					r.data.Empresa,
					r.data.NumeroGuia,
					r.data.TamanoHoja,
					r.data.Orientacion,
					r.data.RemitenteNombre,
					r.data.DestinatarioNombre), window)
		},
		window)

	saveDialog.SetFileName(defaultName)
	saveDialog.SetFilter(storage.NewExtensionFileFilter([]string{".pdf"}))
	saveDialog.Show()
}

func (r *RotuloGenerator) createProfessionalPDF() ([]byte, error) {
	// Obtener dimensiones según tamaño y orientación
	paperSize, ok := paperSizes[r.data.TamanoHoja]
	if !ok {
		paperSize = paperSizes["A4"] // Default
	}

	// Determinar orientación
	orientation := "P" // Portrait (vertical)
	width := paperSize.Width
	height := paperSize.Height

	if r.data.Orientacion == "Horizontal" {
		orientation = "L" // Landscape (horizontal)
		width, height = height, width
	}

	// Crear PDF con gofpdf
	pdf := gofpdf.New(orientation, "mm", r.data.TamanoHoja, "")

	// Intentar cargar fuentes UTF-8, si no existen usar Arial
	fontFamily := "Arial"
	if _, err := os.Stat("fonts/DejaVuSans.ttf"); err == nil {
		pdf.AddUTF8Font("DejaVu", "", "fonts/DejaVuSans.ttf")
		pdf.AddUTF8Font("DejaVu", "B", "fonts/DejaVuSans-Bold.ttf")
		fontFamily = "DejaVu"
	}

	pdf.AddPage()

	// Obtener datos de la empresa
	empresaData := empresasData[r.data.Empresa]

	// Calcular factor de escala basado en el tamaño
	scale := 1.0
	if r.data.TamanoHoja == "A5" {
		scale = 0.7
	} else if r.data.TamanoHoja == "Carta" {
		scale = 1.03
	}

	// Configurar colores corporativos
	pdf.SetFillColor(empresaData.Color.R, empresaData.Color.G, empresaData.Color.B)
	pdf.SetTextColor(255, 255, 255)

	// HEADER - Banda superior con color corporativo
	headerHeight := 20.0 * scale
	pdf.Rect(0, 0, width, headerHeight, "F")

	// Logo (si existe)
	logoPath := zettacomLogo
	if r.data.Empresa == "COMSITEC" {
		logoPath = comsitecLogo
	}

	if _, err := os.Stat(logoPath); err == nil {
		logoWidth := 25.0 * scale
		logoHeight := 12.0 * scale
		pdf.Image(logoPath, 5*scale, 4*scale, logoWidth, logoHeight, false, "", 0, "")
	}

	// Título de la empresa
	pdf.SetFont(fontFamily, "B", 14*scale)
	pdf.SetXY(35*scale, 6*scale)
	pdf.Cell(80*scale, 8*scale, empresaData.Nombre)

	// Número de tracking prominente
	pdf.SetFont(fontFamily, "B", 12*scale)
	pdf.SetXY(width-70*scale, 6*scale)
	pdf.Cell(60*scale, 8*scale, "TRACKING: "+r.data.NumeroGuia)

	// Resetear color de texto
	pdf.SetTextColor(0, 0, 0)

	// Posición inicial después del header
	currentY := headerHeight + 5*scale

	// SECCIÓN FROM y TO en la misma línea
	sectionWidth := (width - 15*scale) / 2

	// FROM (Remitente)
	pdf.SetFont(fontFamily, "B", 10*scale)
	pdf.SetXY(5*scale, currentY)
	pdf.SetFillColor(240, 240, 240)
	pdf.Rect(5*scale, currentY, sectionWidth, 4*scale, "F")
	pdf.Cell(sectionWidth, 4*scale, "FROM / REMITENTE")

	pdf.SetFont(fontFamily, "", 8*scale)
	pdf.SetXY(5*scale, currentY+6*scale)

	// Texto del remitente en líneas controladas
	fromText := fmt.Sprintf("%s", r.data.RemitenteNombre)
	pdf.Cell(sectionWidth, 3*scale, fromText)
	pdf.SetXY(5*scale, currentY+10*scale)

	// Dirección del remitente (máximo 2 líneas)
	fromAddr := strings.ReplaceAll(r.data.RemitenteDireccion, "\n", " ")
	if len(fromAddr) > 40 {
		fromAddr = fromAddr[:40] + "..."
	}
	pdf.Cell(sectionWidth, 3*scale, fromAddr)
	pdf.SetXY(5*scale, currentY+14*scale)
	pdf.Cell(sectionWidth, 3*scale, "Tel: "+r.data.RemitenteTelefono)

	// TO (Destinatario)
	toX := 5*scale + sectionWidth + 5*scale
	pdf.SetFont(fontFamily, "B", 10*scale)
	pdf.SetXY(toX, currentY)
	pdf.SetFillColor(240, 240, 240)
	pdf.Rect(toX, currentY, sectionWidth, 4*scale, "F")
	pdf.Cell(sectionWidth, 4*scale, "TO / DESTINATARIO")

	pdf.SetFont(fontFamily, "", 8*scale)
	pdf.SetXY(toX, currentY+6*scale)

	// Texto del destinatario
	toText := fmt.Sprintf("%s", r.data.DestinatarioNombre)
	pdf.Cell(sectionWidth, 3*scale, toText)
	pdf.SetXY(toX, currentY+10*scale)

	// Dirección del destinatario (máximo 2 líneas)
	toAddr := strings.ReplaceAll(r.data.DestinatarioDireccion, "\n", " ")
	if len(toAddr) > 40 {
		toAddr = toAddr[:40] + "..."
	}
	pdf.Cell(sectionWidth, 3*scale, toAddr)
	pdf.SetXY(toX, currentY+14*scale)
	pdf.Cell(sectionWidth, 3*scale, "Tel: "+r.data.DestinatarioTelefono)

	// Actualizar posición Y
	currentY += 25 * scale

	// INFORMACIÓN DEL ENVÍO
	pdf.SetFont(fontFamily, "B", 10*scale)
	pdf.SetXY(5*scale, currentY)
	pdf.SetFillColor(240, 240, 240)
	pdf.Rect(5*scale, currentY, width-10*scale, 4*scale, "F")
	pdf.Cell(width-10*scale, 4*scale, "DETALLES DEL ENVIO / SHIPMENT DETAILS")

	pdf.SetFont(fontFamily, "", 8*scale)
	currentY += 6 * scale

	// Detalles en líneas controladas
	pdf.SetXY(5*scale, currentY)
	pdf.Cell(width-10*scale, 3*scale, fmt.Sprintf("Fecha/Date: %s", r.data.FechaEnvio.Format("02/01/2006 15:04")))
	currentY += 4 * scale

	if r.data.Peso != "" {
		pdf.SetXY(5*scale, currentY)
		pdf.Cell(width-10*scale, 3*scale, fmt.Sprintf("Peso/Weight: %s", r.data.Peso))
		currentY += 4 * scale
	}

	if r.data.Observaciones != "" {
		pdf.SetXY(5*scale, currentY)
		obsText := r.data.Observaciones
		if len(obsText) > 60 {
			obsText = obsText[:60] + "..."
		}
		pdf.Cell(width-10*scale, 3*scale, fmt.Sprintf("Observaciones/Notes: %s", obsText))
		currentY += 4 * scale
	}

	pdf.SetXY(5*scale, currentY)
	pdf.Cell(width-10*scale, 3*scale, fmt.Sprintf("Servicio/Service: Express | Tamaño/Size: %s - %s", r.data.TamanoHoja, r.data.Orientacion))
	currentY += 8 * scale

	// CÓDIGO DE BARRAS
	pdf.SetFont("Arial", "B", 8*scale) // Usar Arial para el código de barras
	pdf.SetXY(5*scale, currentY)
	pdf.Cell(width-8*scale, 6*scale, "TRACKING NUMBER")
	currentY += 8 * scale

	// Código de barras simplificado con líneas
	pdf.SetFillColor(0, 0, 0) // Negro para las barras
	barWidth := 1.0 * scale
	barHeight := 12.0 * scale
	barSpacing := 2.0 * scale

	// Calcular número de barras que caben
	availableWidth := width - 20*scale
	numBars := int(availableWidth / barSpacing)

	startX := 10 * scale
	for i := 0; i < numBars; i++ {
		// Patrón simple: barra cada 3 posiciones
		if i%3 == 0 || i%7 == 0 {
			pdf.Rect(startX+float64(i)*barSpacing, currentY, barWidth, barHeight, "F")
		}
	}

	currentY += barHeight + 3*scale

	// Número debajo del código de barras
	pdf.SetFont("Arial", "", 10*scale)
	pdf.SetXY(5*scale, currentY)
	pdf.Cell(width-10*scale, 4*scale, r.data.NumeroGuia)
	currentY += 8 * scale

	// Calcular espacio restante
	remainingHeight := height - currentY - 15*scale // Reservar espacio para footer

	// QR CODE (solo para COMSITEC y si hay espacio)
	if empresaData.NeedQR && remainingHeight >= 35*scale {
		qrSize := 25.0 * scale
		qrX := width - qrSize - 5*scale
		qrY := currentY

		qrData := "https://www.comsitec.tech" + r.data.NumeroGuia
		qrCode, err := qrcode.Encode(qrData, qrcode.Medium, 256)
		if err == nil {
			qrPath := "temp_qr.png"
			err = ioutil.WriteFile(qrPath, qrCode, 0644)
			if err == nil {
				pdf.Image(qrPath, qrX, qrY, qrSize, qrSize, false, "", 0, "")
				os.Remove(qrPath)

				pdf.SetFont(fontFamily, "", 6*scale)
				pdf.SetXY(qrX, qrY+qrSize+2*scale)
				pdf.Cell(qrSize, 2*scale, "Escanea para tracking")
			}
		}
	}

	// ÁREA DE FIRMA
	signatureWidth := 70.0 * scale
	signatureHeight := 15.0 * scale
	signatureY := height - 25*scale

	pdf.SetFont(fontFamily, "B", 8*scale)
	pdf.SetXY(5*scale, signatureY-5*scale)
	pdf.Cell(signatureWidth, 3*scale, "FIRMA DESTINATARIO / RECIPIENT SIGNATURE")

	pdf.Rect(5*scale, signatureY, signatureWidth, signatureHeight, "D")

	pdf.SetXY(5*scale, signatureY+signatureHeight+2*scale)
	pdf.SetFont(fontFamily, "", 6*scale)
	pdf.Cell(signatureWidth, 2*scale, "Fecha/Date: _______________")

	// INFORMACIÓN LEGAL/FOOTER

	// INFORMACIÓN LEGAL/FOOTER
	footerY := height - 10*scale
	pdf.SetFont(fontFamily, "", 7*scale)
	pdf.SetXY(10*scale, footerY)
	pdf.MultiCell(width-20*scale, 3*scale, fmt.Sprintf(
		"%s - %s\n"+
			"Este documento constituye comprobante de envío. Conserve para reclamos.\n"+
			"This document constitutes proof of shipment. Keep for claims.\n"+
			"Generado automáticamente el %s",
		empresaData.Nombre,
		empresaData.Direccion,
		time.Now().Format("02/01/2006 15:04")), "", "", false)

	// Usar bytes.Buffer para capturar el output
	var buf bytes.Buffer
	err := pdf.Output(&buf)
	if err != nil {
		return nil, fmt.Errorf("error generando PDF: %v", err)
	}

	return buf.Bytes(), nil
}

func (r *RotuloGenerator) updateLogoPreview(empresa string) {
	logoPath := zettacomLogo
	if empresa == "COMSITEC" {
		logoPath = comsitecLogo
	}

	if _, err := os.Stat(logoPath); os.IsNotExist(err) {
		r.logoPreview.Resource = nil
		r.logoPreview.Refresh()
		return
	}

	r.logoPreview.File = logoPath
	r.logoPreview.Refresh()
}

func (r *RotuloGenerator) updatePreview() {
	if r.preview == nil {
		return
	}

	if r.data.NumeroGuia == "" {
		if r.data.Empresa != "" {
			r.data.NumeroGuia = fmt.Sprintf("%s%d", r.data.Empresa[:3], time.Now().Unix()%1000000)
		} else {
			r.data.NumeroGuia = fmt.Sprintf("GEN%d", time.Now().Unix()%1000000)
		}
	}

	empresaData := empresasData[r.data.Empresa]
	showQR := empresaData.NeedQR

	preview := fmt.Sprintf(`# 🏷️ RÓTULO PROFESIONAL - %s

---

## 📤 FROM / REMITENTE
**%s**
%s
📞 %s

---

## 📥 TO / DESTINATARIO  
**%s**
%s
📞 %s

---

## 📦 DETALLES DEL ENVÍO
- **🔢 Tracking:** %s
- **📅 Fecha:** %s
- **📏 Tamaño:** %s - %s`,
		r.data.Empresa,
		getValueOrDefault(r.data.RemitenteNombre, "[Nombre del remitente]"),
		getValueOrDefault(r.data.RemitenteDireccion, "[Dirección del remitente]"),
		getValueOrDefault(r.data.RemitenteTelefono, "[Teléfono del remitente]"),
		getValueOrDefault(r.data.DestinatarioNombre, "[Nombre del destinatario]"),
		getValueOrDefault(r.data.DestinatarioDireccion, "[Dirección del destinatario]"),
		getValueOrDefault(r.data.DestinatarioTelefono, "[Teléfono del destinatario]"),
		r.data.NumeroGuia,
		time.Now().Format("02/01/2006 15:04"),
		r.data.TamanoHoja,
		r.data.Orientacion,
	)

	if r.data.Peso != "" {
		preview += fmt.Sprintf("\n- **⚖️ Peso:** %s", r.data.Peso)
	}

	if r.data.Observaciones != "" {
		preview += fmt.Sprintf("\n- **📝 Observaciones:** %s", r.data.Observaciones)
	}

	preview += "\n\n---\n\n## ✨ CARACTERÍSTICAS PROFESIONALES\n"
	preview += "✅ Logo corporativo en header\n"
	preview += "✅ Código de barras para tracking\n"
	preview += "✅ Diseño adaptado al tamaño seleccionado\n"
	preview += "✅ Soporte para caracteres especiales (ñ, á, é, etc.)\n"
	preview += "✅ Todo el contenido en una sola página\n"

	if showQR {
		preview += "✅ QR code para tracking online\n"
	}

	preview += "\n---\n*Rótulo profesional generado automáticamente*"

	r.preview.ParseMarkdown(preview)
}

func getValueOrDefault(value, defaultValue string) string {
	if strings.TrimSpace(value) == "" {
		return defaultValue
	}
	return value
}

func (r *RotuloGenerator) printRotulo(window fyne.Window) {
	if r.data.RemitenteNombre == "" || r.data.DestinatarioNombre == "" {
		dialog.ShowError(fmt.Errorf("debes completar al menos el nombre del remitente y destinatario"), window)
		return
	}

	printerOptions := []string{"HP LaserJet Pro", "Epson L3150", "Brother DCP-T510W", "Canon PIXMA", "Impresora predeterminada"}

	printerSelect := widget.NewSelect(printerOptions, nil)
	printerSelect.SetSelected("Impresora predeterminada")

	colorCheck := widget.NewCheck("Imprimir en color", nil)
	colorCheck.SetChecked(true)
	qualityCheck := widget.NewCheck("Alta calidad", nil)
	qualityCheck.SetChecked(true)

	content := container.NewVBox(
		widget.NewLabel("Selecciona la impresora:"),
		printerSelect,
		widget.NewSeparator(),
		widget.NewLabel("Configuración:"),
		colorCheck,
		qualityCheck,
		widget.NewSeparator(),
		widget.NewLabel(fmt.Sprintf("📄 Tamaño: %s - %s", r.data.TamanoHoja, r.data.Orientacion)),
		widget.NewLabel("🎨 Se recomienda impresión en color para mejor resultado"),
	)

	printerDialog := dialog.NewCustomConfirm("Imprimir Rótulo", "Imprimir", "Cancelar", content,
		func(confirmed bool) {
			if confirmed {
				selectedPrinter := printerSelect.Selected
				dialog.ShowInformation("✅ Impresión Enviada",
					fmt.Sprintf("Rótulo profesional enviado a: %s\n\n"+
						"🏢 Empresa: %s\n"+
						"📦 Tracking: %s\n"+
						"📏 Tamaño: %s - %s\n"+
						"🎨 Color: %v\n"+
						"⭐ Alta calidad: %v\n\n"+
						"El rótulo incluye logo, código de barras y diseño profesional.",
						selectedPrinter,
						r.data.Empresa,
						r.data.NumeroGuia,
						r.data.TamanoHoja,
						r.data.Orientacion,
						colorCheck.Checked,
						qualityCheck.Checked), window)
			}
		}, window)

	printerDialog.Show()
}

func (r *RotuloGenerator) clearFields() {
	for _, entry := range r.inputs {
		entry.SetText("")
	}
	r.data = &RotuloData{
		TamanoHoja:  "A4",
		Orientacion: "Vertical",
		FechaEnvio:  time.Now(),
	}
	r.empresaCheck.SetSelected("ZETTACOM")
	r.data.Empresa = "ZETTACOM"
	r.tamanoHoja.SetSelected("A4")
	r.orientacion.SetSelected("Vertical")
	r.updateLogoPreview("ZETTACOM")
	r.updatePreview()
}

func (r *RotuloGenerator) fillTestData() {
	r.empresaCheck.SetSelected("COMSITEC")
	r.data.Empresa = "COMSITEC"
	r.updateLogoPreview("COMSITEC")

	r.inputs["destinatarioNombre"].SetText("María González López")
	r.inputs["destinatarioDireccion"].SetText("Jr. Los Olivos 456\nMiraflores, Lima 15074\nPerú")
	r.inputs["destinatarioTelefono"].SetText("+51 888 777 666")
	r.inputs["peso"].SetText("2.5 kg")
	r.inputs["observaciones"].SetText("FRÁGIL - Manejar con cuidado")
	r.inputs["numeroGuia"].SetText("COM123456")
	r.tamanoHoja.SetSelected("A4")
	r.orientacion.SetSelected("Vertical")
}

// Funciones del notepad (mantenidas igual)...

func (n *NotePad) createPersonalTab(window fyne.Window) *fyne.Container {
	n.multiLine = widget.NewMultiLineEntry()
	n.multiLine.Wrapping = fyne.TextWrapOff
	n.multiLine.Resize(fyne.NewSize(600, 300))

	n.multiLine.OnChanged = func(content string) {
		n.lastContent = content
		n.lastSaveTime = time.Now()
		n.lastUserEdit = time.Now()
		if n.statusLabel != nil {
			n.statusLabel.SetText("Estado: Modificado (guardado automático)")
		}
	}

	n.loadContent()

	scroll := container.NewScroll(n.multiLine)
	scroll.SetMinSize(fyne.NewSize(600, 300))

	n.statusLabel = widget.NewLabel("Estado: Listo")
	timeLabel := widget.NewLabel(fmt.Sprintf("Última actualización: %s", time.Now().Format("15:04:05")))

	saveButton := widget.NewButton("💾 Guardar Ahora", func() {
		n.saveContent()
		n.statusLabel.SetText("Estado: Guardado manualmente")
		go func() {
			time.Sleep(2 * time.Second)
			n.statusLabel.SetText("Estado: Listo")
		}()
	})

	reloadButton := widget.NewButton("🔄 Recargar", func() {
		n.loadContent()
		n.statusLabel.SetText("Estado: Recargado desde archivo")
		go func() {
			time.Sleep(2 * time.Second)
			n.statusLabel.SetText("Estado: Listo")
		}()
	})

	clearButton := widget.NewButton("🗑️ Limpiar", func() {
		dialog.ShowConfirm("Confirmar", "¿Estás seguro de que quieres limpiar todo el contenido?", func(confirmed bool) {
			if confirmed {
				n.multiLine.SetText("")
				n.statusLabel.SetText("Estado: Contenido limpiado")
			}
		}, window)
	})

	autoUpdateInfo := widget.NewRichTextFromMarkdown(`
**Actualización Automática de Hora:**

La hora se actualiza automáticamente cada segundo en el texto.
- Detecta patrones como "11:24", "17:11", etc.
- Solo actualiza si no has editado recientemente (2 segundos de pausa)
- Preserva la posición del cursor
- No interfiere con tu escritura

**Ejemplo:**
Si escribes "REPOSICION 15:30 JRIOS", la hora se actualizará automáticamente a la hora actual.
`)
	autoUpdateInfo.Wrapping = fyne.TextWrapWord

	infoScroll := container.NewScroll(autoUpdateInfo)
	infoScroll.SetMinSize(fyne.NewSize(300, 200))

	go n.startTimeUpdates(timeLabel)
	go n.startAutoSave()

	editorCard := widget.NewCard("📝 Editor de Texto", "",
		container.NewVBox(
			container.NewHBox(saveButton, reloadButton, clearButton),
			scroll,
		),
	)

	infoCard := widget.NewCard("ℹ️ Actualización Automática", "", infoScroll)

	statusCard := widget.NewCard("📊 Estado", "",
		container.NewVBox(n.statusLabel, timeLabel),
	)

	return container.NewVBox(
		widget.NewLabel("Bloc de notas con fecha actualizada"),
		container.NewHBox(
			container.NewVBox(editorCard, statusCard),
			infoCard,
		),
	)
}

func (n *NotePad) startTimeUpdates(timeLabel *widget.Label) {
	ticker := time.NewTicker(time.Second)
	defer ticker.Stop()

	for range ticker.C {
		now := time.Now()
		currentTime := now.Format("15:04")
		content := n.multiLine.Text

		timeLabel.SetText(fmt.Sprintf("Última actualización: %s", now.Format("15:04:05")))

		if time.Since(n.lastUserEdit) < 2*time.Second {
			continue
		}

		timeRegex := regexp.MustCompile(`\b\d{1,2}:\d{2}\b`)

		if timeRegex.MatchString(content) {
			newContent := timeRegex.ReplaceAllString(content, currentTime)

			if newContent != content {
				cursorRow := n.multiLine.CursorRow
				cursorCol := n.multiLine.CursorColumn

				n.multiLine.SetText(newContent)

				n.multiLine.CursorRow = cursorRow
				n.multiLine.CursorColumn = cursorCol

				n.lastContent = newContent
			}
		}
	}
}

func (n *NotePad) startAutoSave() {
	ticker := time.NewTicker(autoSaveInterval)
	defer ticker.Stop()

	for range ticker.C {
		if time.Since(n.lastSaveTime) >= 2*time.Second && n.lastContent != "" {
			n.saveContent()
		}
	}
}

func (n *NotePad) saveContent() {
	content := n.multiLine.Text
	if content == "" {
		return
	}

	dir := filepath.Dir(saveFile)
	if dir != "." {
		os.MkdirAll(dir, 0755)
	}

	timestamp := time.Now().Format("2006-01-02 15:04:05")
	contentWithTimestamp := fmt.Sprintf("# Guardado: %s\n%s", timestamp, content)

	err := ioutil.WriteFile(saveFile, []byte(contentWithTimestamp), 0644)
	if err != nil {
		log.Printf("Error guardando archivo: %v", err)
	}
}

func (n *NotePad) loadContent() {
	if _, err := os.Stat(saveFile); os.IsNotExist(err) {
		defaultContent := `***********LISTA REPOSICIÓN*********
......9999 REPOSICION 15:04 MGAVINO
......9999 REPOSICION 15:04 JRIOS
......9999 REPOSICION 15:04 BTAIPE
......9999 REPOSICION 15:04 MQUINTANA

**************ZETTACOM**********
......0154 LGARCIA 15:04 MGAVINO
......0154 LGARCIA 15:04 JRIOS
......0083 JVILCATOMA 15:04 MGAVINO
......0017 NCRISOSTOMO 15:04 JRIOS

# Las horas se actualizan automáticamente cada segundo
# Puedes editar el texto libremente
# Solo espera 2 segundos después de escribir para que se actualice la hora`

		n.multiLine.SetText(defaultContent)
		n.lastContent = defaultContent
		return
	}

	data, err := ioutil.ReadFile(saveFile)
	if err != nil {
		log.Printf("Error cargando archivo: %v", err)
		return
	}

	content := string(data)
	lines := strings.Split(content, "\n")
	if len(lines) > 0 && strings.HasPrefix(lines[0], "# Guardado:") {
		content = strings.Join(lines[1:], "\n")
	}

	n.multiLine.SetText(content)
	n.lastContent = content
}

func globalEscapeListener(statusLabel *widget.Label) {
	fmt.Println("Listener global de ESC activado.")
	hook.Register(hook.KeyDown, []string{"esc"}, func(e hook.Event) {
		select {
		case <-cancel:
		default:
			close(cancel)
			if statusLabel != nil {
				statusLabel.SetText("Estado: Cancelado con ESC.")
			}
			fmt.Println("Escape presionado.")
		}
	})

	s := hook.Start()
	<-hook.Process(s)
}

func autocopiar(rawSeries string, date string, delay time.Duration, countdown int, statusLabel, copiedCounter *widget.Label) {
	time.Sleep(3 * time.Second)

	series := strings.Fields(rawSeries)
	total := len(series)
	copied := 0

	for i := countdown; i > 0; i-- {
		statusLabel.SetText(fmt.Sprintf("Comenzando en %d...", i))
		select {
		case <-cancel:
			return
		default:
		}
		time.Sleep(time.Second)
	}

	statusLabel.SetText("Copiando...")

	for _, s := range series {
		select {
		case <-cancel:
			statusLabel.SetText("Estado: Cancelado.")
			return
		default:
		}
		robotgo.TypeStrDelay(s, 2)
		time.Sleep(delay)

		robotgo.KeyTap("tab")
		time.Sleep(delay)

		robotgo.TypeStrDelay(date, 2)
		time.Sleep(delay)

		robotgo.KeyTap("down")
		time.Sleep(60 * time.Millisecond)

		copied++
		copiedCounter.SetText(fmt.Sprintf("Copiadas: %d / %d", copied, total))
	}

	statusLabel.SetText("Estado: Finalizado correctamente.")
}
