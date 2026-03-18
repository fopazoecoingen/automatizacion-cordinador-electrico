using System.Diagnostics;
using GeneradorInformeElectrico.Models;
using GeneradorInformeElectrico.Services;

namespace GeneradorInformeElectrico;

public partial class MainForm : Form
{
    private ComboBox _mesCombo = null!;
    private ComboBox _empresaCombo = null!;
    private ComboBox _barraCombo = null!;
    private TextBox _carpetaDatosBox = null!;
    private TextBox _plantillaBox = null!;
    private TextBox _destinoBox = null!;
    private ProgressBar _progressBar = null!;
    private Label _progressLabel = null!;
    private Button _createBtn = null!;
    private NumericUpDown _anyoNum = null!;
    private bool _procesando;

    private static readonly string[] MesesNombres =
    {
        "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
        "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
    };

    public MainForm()
    {
        InitializeComponent();
        CargarUltimosDatos();
    }

    private void InitializeComponent()
    {
        Text = "Generación de Informe Eléctrico";
        Size = new Size(920, 760);
        MinimumSize = new Size(700, 600);
        BackColor = Color.FromArgb(240, 244, 248);
        Font = new Font("Segoe UI", 9);

        var mainPanel = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(24),
            BackColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle
        };

        var header = new Panel
        {
            Height = 60,
            Dock = DockStyle.Top,
            BackColor = Color.FromArgb(26, 54, 93)
        };
        var titleLbl = new Label
        {
            Text = "Generación de Informe Eléctrico",
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 18, FontStyle.Bold),
            AutoSize = true,
            Location = new Point(24, 18)
        };
        header.Controls.Add(titleLbl);

        var y = 20;

        var periodoLbl = new Label { Text = "Período", Font = new Font("Segoe UI", 11, FontStyle.Bold), Location = new Point(32, y) };
        mainPanel.Controls.Add(periodoLbl);
        y += 28;

        _anyoNum = new NumericUpDown
        {
            Minimum = 2020,
            Maximum = 2030,
            Value = DateTime.Now.Year,
            Width = 80,
            Location = new Point(32, y)
        };
        _mesCombo = new ComboBox
        {
            Width = 150,
            Location = new Point(120, y),
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        _mesCombo.Items.AddRange(MesesNombres);
        _mesCombo.SelectedIndex = DateTime.Now.Month - 1;
        mainPanel.Controls.Add(_anyoNum);
        mainPanel.Controls.Add(_mesCombo);
        y += 45;

        var empLbl = new Label { Text = "Empresa", Font = new Font("Segoe UI", 11, FontStyle.Bold), Location = new Point(32, y) };
        mainPanel.Controls.Add(empLbl);
        y += 28;
        _empresaCombo = new ComboBox
        {
            Width = 300,
            Location = new Point(32, y),
            DropDownStyle = ComboBoxStyle.DropDown
        };
        _empresaCombo.TextChanged += (_, _) => ActualizarBarras();
        mainPanel.Controls.Add(_empresaCombo);
        y += 45;

        var barraLbl = new Label { Text = "Barra", Font = new Font("Segoe UI", 11, FontStyle.Bold), Location = new Point(32, y) };
        mainPanel.Controls.Add(barraLbl);
        y += 28;
        _barraCombo = new ComboBox { Width = 300, Location = new Point(32, y), DropDownStyle = ComboBoxStyle.DropDown };
        mainPanel.Controls.Add(_barraCombo);
        y += 55;

        var carpetaDatosLbl = new Label { Text = "Carpeta donde se guardará la base de datos (archivos Balance, Anexos PLABACOM)", Font = new Font("Segoe UI", 11, FontStyle.Bold), Location = new Point(32, y), AutoSize = true };
        mainPanel.Controls.Add(carpetaDatosLbl);
        y += 28;
        _carpetaDatosBox = new TextBox { Width = 500, Location = new Point(32, y), ReadOnly = true };
        var btnCarpetaDatos = new Button { Text = "Examinar", Location = new Point(540, y - 2), Width = 80 };
        btnCarpetaDatos.Click += (_, _) =>
        {
            using var dlg = new FolderBrowserDialog
            {
                Description = "Seleccione la carpeta donde se descargarán y guardarán los archivos Excel (Balance, Anexos PLABACOM)",
                UseDescriptionForTitle = true,
                SelectedPath = _carpetaDatosBox.Text.Trim()
            };
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                _carpetaDatosBox.Text = dlg.SelectedPath;
                ConfigService.GuardarCarpetaBaseDatos(dlg.SelectedPath);
            }
        };
        mainPanel.Controls.Add(_carpetaDatosBox);
        mainPanel.Controls.Add(btnCarpetaDatos);
        y += 45;

        var plantillaLbl = new Label { Text = "Plantilla Excel del cliente (archivo base del informe a completar)", Font = new Font("Segoe UI", 11, FontStyle.Bold), Location = new Point(32, y), AutoSize = true, MaximumSize = new Size(550, 0) };
        mainPanel.Controls.Add(plantillaLbl);
        y += 28;
        _plantillaBox = new TextBox { Width = 500, Location = new Point(32, y), ReadOnly = true };
        var btnPlantilla = new Button { Text = "Examinar", Location = new Point(540, y - 2), Width = 80 };
        btnPlantilla.Click += (_, _) =>
        {
            using var dlg = new OpenFileDialog
            {
                Filter = "Excel|*.xlsx;*.xlsm",
                Title = "Seleccionar plantilla"
            };
            if (dlg.ShowDialog() == DialogResult.OK)
                _plantillaBox.Text = dlg.FileName;
        };
        mainPanel.Controls.Add(_plantillaBox);
        mainPanel.Controls.Add(btnPlantilla);
        y += 45;

        var destLbl = new Label { Text = "Ruta y nombre del archivo informe final donde se guardará el resultado", Font = new Font("Segoe UI", 11, FontStyle.Bold), Location = new Point(32, y), AutoSize = true, MaximumSize = new Size(550, 0) };
        mainPanel.Controls.Add(destLbl);
        y += 28;
        _destinoBox = new TextBox { Width = 500, Location = new Point(32, y), ReadOnly = true };
        var btnDestino = new Button { Text = "Examinar", Location = new Point(540, y - 2), Width = 80 };
        btnDestino.Click += (_, _) =>
        {
            using var dlg = new SaveFileDialog
            {
                Filter = "Excel|*.xlsx;*.xlsm",
                Title = "Guardar informe como"
            };
            if (dlg.ShowDialog() == DialogResult.OK)
                _destinoBox.Text = dlg.FileName;
        };
        mainPanel.Controls.Add(_destinoBox);
        mainPanel.Controls.Add(btnDestino);
        y += 55;

        _progressBar = new ProgressBar { Width = 600, Height = 25, Location = new Point(32, y) };
        _progressLabel = new Label { Text = "", Location = new Point(32, y + 30), AutoSize = true, ForeColor = Color.Gray };
        mainPanel.Controls.Add(_progressBar);
        mainPanel.Controls.Add(_progressLabel);
        y += 70;

        _createBtn = new Button
        {
            Text = "  Crear Informe  ",
            BackColor = Color.FromArgb(43, 108, 176),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            Size = new Size(180, 45),
            Location = new Point(32, y),
            FlatStyle = FlatStyle.Flat
        };
        _createBtn.Click += CreateBtn_Click;
        mainPanel.Controls.Add(_createBtn);

        var container = new Panel { Dock = DockStyle.Fill, Padding = new Padding(24), BackColor = Color.FromArgb(240, 244, 248) };
        container.Controls.Add(mainPanel);
        mainPanel.Dock = DockStyle.Fill;

        Controls.Add(container);
        Controls.Add(header);

        CargarEmpresas();
    }

    private void CargarEmpresas()
    {
        var config = ConfigService.CargarConfigEmpresas();
        _empresaCombo.Items.Clear();
        _empresaCombo.Items.Add("");
        foreach (var emp in config.Empresas)
        {
            if (!string.IsNullOrEmpty(emp.NombreEmpresa))
                _empresaCombo.Items.Add(emp.NombreEmpresa);
        }
    }

    private void ActualizarBarras()
    {
        var config = ConfigService.CargarConfigEmpresas();
        var emp = config.Empresas.FirstOrDefault(e => string.Equals(e.NombreEmpresa, _empresaCombo.Text, StringComparison.OrdinalIgnoreCase));
        _barraCombo.Items.Clear();
        _barraCombo.Items.Add("");
        if (emp?.Barras != null)
            foreach (var b in emp.Barras)
                _barraCombo.Items.Add(b);
    }

    private void CargarUltimosDatos()
    {
        _carpetaDatosBox.Text = ConfigService.GetCarpetaBaseDatos();
        var datos = ConfigService.CargarUltimosDatos();
        if (datos == null) return;
        if (datos.Anyo.HasValue) _anyoNum.Value = datos.Anyo.Value;
        if (datos.Mes.HasValue && datos.Mes >= 1 && datos.Mes <= 12) _mesCombo.SelectedIndex = datos.Mes.Value - 1;
        if (!string.IsNullOrEmpty(datos.Empresa)) _empresaCombo.Text = datos.Empresa;
        if (!string.IsNullOrEmpty(datos.Barra)) _barraCombo.Text = datos.Barra;
        if (!string.IsNullOrEmpty(datos.Plantilla)) _plantillaBox.Text = datos.Plantilla;
        if (!string.IsNullOrEmpty(datos.Destino)) _destinoBox.Text = datos.Destino;
    }

    private void CreateBtn_Click(object? sender, EventArgs e)
    {
        if (_procesando) return;
        var plantilla = _plantillaBox.Text.Trim();
        var destino = _destinoBox.Text.Trim();
        if (string.IsNullOrEmpty(plantilla))
        {
            MessageBox.Show("Por favor seleccione su plantilla Excel.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }
        if (string.IsNullOrEmpty(destino))
        {
            MessageBox.Show("Por favor seleccione una ruta de destino para el informe.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }
        if (!File.Exists(plantilla))
        {
            MessageBox.Show("La plantilla no existe.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        var anyo = (int)_anyoNum.Value;
        var mes = _mesCombo.SelectedIndex + 1;
        var nombreEmpresa = _empresaCombo.Text.Trim();
        var nombreBarra = _barraCombo.Text.Trim();

        ConfigService.GuardarUltimosDatos(new UltimosDatos
        {
            Anyo = anyo,
            Mes = mes,
            Empresa = nombreEmpresa,
            Barra = nombreBarra,
            Plantilla = plantilla,
            Destino = destino
        });

        _procesando = true;
        _createBtn.Enabled = false;
        _createBtn.Text = "Procesando...";
        _progressBar.Value = 0;
        _progressLabel.Text = "Iniciando...";

        LogHelper.Log($"Iniciando proceso: año={anyo} mes={mes} empresa={nombreEmpresa} barra={nombreBarra}");
        LogHelper.Log($"Plantilla: {plantilla}");
        LogHelper.Log($"Destino: {destino}");

        Task.Run(async () =>
        {
            try
            {
                await ProcesarMesAsync(anyo, mes, plantilla, destino, nombreEmpresa, nombreBarra);
                LogHelper.Log("Proceso completado correctamente.");
                Invoke(() =>
                {
                    _progressBar.Value = 100;
                    _progressLabel.Text = "[OK] Proceso completado";
                    MessageBox.Show($"Informe generado exitosamente.\n\nArchivo: {destino}", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                });
            }
            catch (Exception ex)
            {
                LogHelper.LogExcepcion(ex);
                var msg = !string.IsNullOrEmpty(ex.Message) ? ex.Message : ex.ToString();
                if (ex.InnerException != null)
                    msg += $"\n\nDetalle: {ex.InnerException.Message}";
                Invoke(() => MessageBox.Show($"Error: {msg}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error));
            }
            finally
            {
                Invoke(() =>
                {
                    _procesando = false;
                    _createBtn.Enabled = true;
                    _createBtn.Text = "  Crear Informe  ";
                    _progressBar.Value = 0;
                });
            }
        });
    }

    private async Task ProcesarMesAsync(int anyo, int mes, string plantilla, string destino, string nombreEmpresa, string nombreBarra)
    {
        var carpetaBd = ConfigService.GetCarpetaBaseDatos();
        var carpetaDescomprimidos = Path.Combine(carpetaBd, "descomprimidos");
        Directory.CreateDirectory(carpetaBd);
        Directory.CreateDirectory(carpetaDescomprimidos);

        LogHelper.Log("Paso 1: Descargando archivos...");
        ReportarProgreso(5, "Descargando archivos...");
        var tipos = new[] { "energia_resultados", "sscc", "potencia" };
        foreach (var tipo in tipos)
        {
            var (rutaZip, rutaDes, codError) = await DescargaService.DescargarYDescomprimirAsync(
                anyo, mes, tipo, carpetaBd, carpetaDescomprimidos,
                new Progress<string>(s => ReportarProgreso(10, s)));
            if (tipo == "energia_resultados" && codError == 403)
                throw new Exception("No se encuentra la información disponible para la descarga para este período.");
        }

        LogHelper.Log("Paso 2: Copiando plantilla...");
        ReportarProgreso(30, "Copiando plantilla...");
        try
        {
            File.Copy(plantilla, destino, overwrite: true);
        }
        catch (IOException io) when (io.Message.Contains("being used"))
        {
            throw new InvalidOperationException(
                $"No se puede escribir en '{Path.GetFileName(destino)}' porque está abierto en Excel u otro programa. " +
                "Cierre el archivo e intente de nuevo.", io);
        }

        LogHelper.Log("Paso 3: Leyendo Balance Valorizado...");
        ReportarProgreso(40, "Leyendo Balance Valorizado...");
        var balanceData = ExcelLecturaService.LeerBalanceValorizado(anyo, mes, carpetaBd);
        if (balanceData == null)
            throw new Exception("No se encontró el archivo Balance para este período.");

        var empresaConfig = ConfigService.CargarConfigEmpresas().Empresas
            .FirstOrDefault(e => string.Equals(e.NombreEmpresa, nombreEmpresa, StringComparison.OrdinalIgnoreCase));

        var filas = balanceData.Filas.AsEnumerable();
        if (!string.IsNullOrEmpty(nombreEmpresa) && balanceData.Filas.Count > 0 && balanceData.Filas[0].ContainsKey("nombre_corto_empresa"))
            filas = filas.Where(f => string.Equals((f.GetValueOrDefault("nombre_corto_empresa")?.ToString() ?? "").Trim(), nombreEmpresa, StringComparison.OrdinalIgnoreCase));
        if (!string.IsNullOrEmpty(nombreBarra) && balanceData.Filas.Count > 0 && balanceData.Filas[0].ContainsKey("barra"))
            filas = filas.Where(f => string.Equals((f.GetValueOrDefault("barra")?.ToString() ?? "").Trim(), nombreBarra, StringComparison.OrdinalIgnoreCase));
        var dfGuardar = filas.ToList();

        var medidorImportacion = empresaConfig?.IMPORTACION_MWh?.Trim();
        var medidoresEnergia = empresaConfig?.GetMedidoresEnergia() ?? new List<string>();
        var conceptoFiltro = empresaConfig?.GetConceptoPotenciaFirme() ?? new List<string> { "Eólica" };

        LogHelper.Log("Paso 4: Calculando valores...");
        ReportarProgreso(50, "Calculando valores...");

        double totalMonetario = ExcelLecturaService.LeerTotalIngresosPotenciaFirme(anyo, mes, carpetaBd, nombreEmpresa, conceptoFiltro) ?? 0;
        if (totalMonetario == 0 && dfGuardar.Count > 0 && dfGuardar[0].ContainsKey("monetario"))
            totalMonetario = dfGuardar.Sum(f => ToDouble(f.GetValueOrDefault("monetario")));

        var totalIt = ExcelLecturaService.LeerIngresosPorIt(anyo, mes, carpetaBd, nombreEmpresa);
        var totalPotencia = ExcelLecturaService.LeerIngresosPorPotencia(anyo, mes, carpetaBd, nombreEmpresa);

        // TOTAL INGRESOS POR POTENCIA FIRME CLP = Pago PSUF + INGRESOS POR IT (fuente: BDef Balance2 + Anexo 02.b)
        totalMonetario += totalIt ?? 0;

        var dfTotalEnergia = dfGuardar;
        if (medidoresEnergia.Count > 0 && dfGuardar.Count > 0 && dfGuardar[0].ContainsKey("nombre_medidor"))
            dfTotalEnergia = dfGuardar.Where(f => medidoresEnergia.Contains((f.GetValueOrDefault("nombre_medidor")?.ToString() ?? "").Trim(), StringComparer.OrdinalIgnoreCase)).ToList();
        double totalEnergia = dfTotalEnergia.Count > 0 ? ToDouble(dfTotalEnergia[0].GetValueOrDefault("monetario")) : 0;

        var nemotecnicoSscc = empresaConfig?.SSCC_NEMOTECNICO?.Trim();
        var filtroSscc = !string.IsNullOrEmpty(nemotecnicoSscc) ? nemotecnicoSscc : nombreEmpresa;
        var totalSscc = !string.IsNullOrEmpty(filtroSscc) ? ExcelLecturaService.LeerTotalIngresosSscc(anyo, mes, filtroSscc, carpetaBd) : null;
        var barraGm = !string.IsNullOrEmpty(nombreBarra) ? nombreBarra : (empresaConfig?.Barras?.FirstOrDefault()?.Trim());
        var totalGmHoldings = ExcelLecturaService.LeerCompraVentaEnergiaGmHoldings(anyo, mes, nombreEmpresa, barraGm, carpetaBd);

        var dfImportacion = dfGuardar;
        if (!string.IsNullOrEmpty(medidorImportacion) && dfGuardar.Count > 0 && dfGuardar[0].ContainsKey("nombre_medidor"))
            dfImportacion = dfGuardar.Where(f => string.Equals((f.GetValueOrDefault("nombre_medidor")?.ToString() ?? "").Trim(), medidorImportacion, StringComparison.OrdinalIgnoreCase)).ToList();
        double? importacionMwh = dfImportacion.Count > 0 && dfImportacion[0].ContainsKey("fisico_kwh")
            ? Math.Abs(dfImportacion.Sum(f => ParseDouble(f.GetValueOrDefault("fisico_kwh")?.ToString()) ?? 0)) / 1000.0
            : null;

        LogHelper.Log($"  Valores: POTENCIA_FIRME={totalMonetario:N0}, IT={totalIt?.ToString("N0") ?? "null"}, POTENCIA={totalPotencia?.ToString("N0") ?? "null"}, SSCC={totalSscc?.ToString("N0") ?? "null"}, GM={totalGmHoldings?.ToString("N0") ?? "null"}, IMP_MWh={importacionMwh?.ToString("N2") ?? "null"}");

        var pares = new List<(string, double)>
        {
            ("TOTAL INGRESOS POR POTENCIA FIRME CLP", totalMonetario)
        };
        if (totalIt.HasValue) pares.Add(("INGRESOS POR IT POTENCIA", totalIt.Value));
        if (totalPotencia.HasValue) pares.Add(("INGRESOS POR POTENCIA", totalPotencia.Value));
        // Valor directo (ParseMonetario lee el string formateado como Python)
        if (totalEnergia != 0) pares.Add(("TOTAL INGRESOS POR ENERGIA CLP", totalEnergia));
        if (!string.IsNullOrEmpty(nombreEmpresa) && totalSscc.HasValue) pares.Add(("TOTAL INGRESOS POR SSCC CLP", totalSscc.Value));
        if (totalGmHoldings.HasValue) pares.Add(("Compra Venta Energia GM Holdings CLP", totalGmHoldings.Value / 1000.0));
        if (importacionMwh.HasValue) pares.Add(("IMPORTACION MWh", importacionMwh.Value));

        LogHelper.Log("Paso 5: Escribiendo en plantilla con Excel COM...");
        ReportarProgreso(80, "Escribiendo en plantilla...");
        ExcelEscrituraService.EscribirTodosEnResultado(destino, anyo, mes, pares);
        LogHelper.Log("Escritura en plantilla finalizada.");
    }

    private void ReportarProgreso(int value, string text)
    {
        if (IsDisposed) return;
        try
        {
            Invoke(() =>
            {
                _progressBar.Value = Math.Min(100, value);
                _progressLabel.Text = text;
            });
        }
        catch { }
    }

    /// <summary>Convierte valor monetario a double sin perder decimales (ParseDouble eliminaba la coma en "325161639,078").</summary>
    private static double ToDouble(object? val)
    {
        if (val is double d) return d;
        if (val is int i) return i;
        if (val is float f) return f;
        if (val is decimal dec) return (double)dec;
        var s = (val?.ToString() ?? "").Trim().Replace("$", "").Replace(" ", "");
        if (string.IsNullOrEmpty(s)) return 0;
        if (s.Contains(",")) s = s.Replace(".", "").Replace(",", ".");
        else if (!s.Contains(".")) s = s.Replace(".", "").Replace(",", ".");
        if (!double.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var v)) return 0;
        if (v >= 1e9 && Math.Abs(v - Math.Floor(v)) < 1e-9) return v / 1000;
        return v;
    }

    private static double? ParseDouble(string? s)
    {
        if (string.IsNullOrWhiteSpace(s)) return null;
        s = s.Replace(",", "").Replace(".", System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator);
        return double.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var v) ? v : null;
    }
}
