using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Security.Cryptography;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using System.Xml.Linq;


namespace FranchiseKPIDashboard
{
    public partial class KpiDashboardForm : Form
    {
        // ------------------------
        // Attributes (Fields)
        // ------------------------
        private List<int> dailySales;
        private string franchiseName = "Demo Franchise";

        // ------------------------
        // Constructors
        // ------------------------
        public KpiDashboardForm()
        {
            InitializeComponent();
            ApplyVisualTheme();
            dailySales = new List<int>();
        }

        private void ApplyVisualTheme()
        {
            // 🎨 Color Palette (60-30-10)
            Color dominantBg = ColorTranslator.FromHtml("#F1F1F1"); // 60% Light Gray
            Color secondaryBg = ColorTranslator.FromHtml("#1D4D4F"); // 30% Deep Green
            Color accentColor = ColorTranslator.FromHtml("#FFA500"); // 10% Harvest Orange
            Color chartAreaBg = Color.White;
            Font headerFont = new Font("Segoe UI", 18, FontStyle.Bold);
            Font labelFont = new Font("Segoe UI", 10, FontStyle.Regular);

            // 1) Form Background
            this.BackColor = dominantBg;

            // 2) Header Label
            var lblHeader = new Label
            {
                Name = "lblHeader",
                Text = "Franchise KPI Dashboard",
                Font = headerFont,
                ForeColor = secondaryBg,
                Dock = DockStyle.Top,
                Height = 50,
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = dominantBg
            };
            this.Controls.Add(lblHeader);
            lblHeader.BringToFront();

            // 3) Chart Panels & Styling
            foreach (var chart in new[] { chartSales, chartRevenue, chartCustomers })
            {
                // wrap each chart in a Panel for padding
                var panel = new Panel
                {
                    BackColor = dominantBg,
                    BorderStyle = BorderStyle.None,
                    Padding = new Padding(10),
                    Width = chart.Width + 20,
                    Height = chart.Height + 20,
                };
                chart.Parent.Controls.Add(panel);
                panel.Location = chart.Location;
                chart.Location = new Point(10, 10);
                panel.Controls.Add(chart);

                // chart styling
                chart.BackColor = dominantBg;
                chart.ChartAreas[0].BackColor = chartAreaBg;
                chart.ChartAreas[0].AxisX.LineColor = secondaryBg;
                chart.ChartAreas[0].AxisY.LineColor = secondaryBg;
                chart.ChartAreas[0].AxisX.LabelStyle.ForeColor = secondaryBg;
                chart.ChartAreas[0].AxisY.LabelStyle.ForeColor = secondaryBg;
                chart.ChartAreas[0].BackSecondaryColor = dominantBg;
                chart.ChartAreas[0].ShadowColor = Color.Gray;
                chart.ChartAreas[0].ShadowOffset = 2;
                chart.Series[0].Color = accentColor;
                chart.Series[0].BorderWidth = 2;
                chart.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;

                // add a title

                chartRevenue.Titles.Clear();
                chartRevenue.Titles.Add("Revenue");
                chartRevenue.Titles[0].Font = new Font("Segoe UI", 9, FontStyle.Bold);
                chartRevenue.Titles[0].Alignment = ContentAlignment.TopCenter;
                chartRevenue.Legends.Clear();
                chartRevenue.BorderlineDashStyle = ChartDashStyle.Solid;
                chartRevenue.BorderlineColor = Color.LightGray;
                chartRevenue.BorderSkin.SkinStyle = BorderSkinStyle.FrameThin2;
                chartRevenue.BackColor = Color.White;
                chartRevenue.ChartAreas[0].BackColor = Color.White;

                chartCustomers.Titles.Clear();
                chartCustomers.Titles.Add("Customers");
                chartCustomers.Titles[0].Font = new Font("Segoe UI", 9, FontStyle.Bold);
                chartCustomers.Titles[0].Alignment = ContentAlignment.TopCenter;
                chartCustomers.Legends.Clear();
                chartCustomers.BackColor = Color.White;
                chartCustomers.ChartAreas[0].BackColor = Color.White;
                chartCustomers.BorderlineDashStyle = ChartDashStyle.Solid;
                chartCustomers.BorderlineColor = Color.LightGray;
                chartCustomers.BorderSkin.SkinStyle = BorderSkinStyle.FrameThin2;

                chartSales.Titles.Clear();
                chartSales.Titles.Add("Sales");
                chartSales.Titles[0].Font = new Font("Segoe UI", 9, FontStyle.Bold);
                chartSales.Titles[0].Alignment = ContentAlignment.TopCenter;
                chartSales.Legends.Clear();
                chartSales.BackColor = Color.White;
                chartSales.ChartAreas[0].BackColor = Color.White;
                chartSales.BorderlineDashStyle = ChartDashStyle.Solid;
                chartSales.BorderlineColor = Color.LightGray;
                chartSales.BorderSkin.SkinStyle = BorderSkinStyle.FrameThin2;
                                             
              
            }

            chartSales.Titles.Clear();
            chartSales.Titles.Add("Sales");

            chartRevenue.Titles.Clear();
            chartRevenue.Titles.Add("Revenue");

            chartCustomers.Titles.Clear();
            chartCustomers.Titles.Add("Customers");

            // Optional formatting:
            chartSales.Titles[0].Font = new Font("Segoe UI", 10, FontStyle.Bold);
            chartRevenue.Titles[0].Font = new Font("Segoe UI", 10, FontStyle.Bold);
            chartCustomers.Titles[0].Font = new Font("Segoe UI", 10, FontStyle.Bold);


            // 4) Button Styling



            foreach (var btn in new[] { btnLoadData, btnExportPDF, btnLoadFromExcel })

            {
                btn.FlatStyle = FlatStyle.Flat;
                btn.FlatAppearance.BorderSize = 0;
                btn.BackColor = accentColor;
                btn.ForeColor = Color.White;
                btn.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                btn.Padding = new Padding(8, 4, 8, 4);
            }

            // 5) Layout Adjustments

            // (Optionally center buttons in a FlowLayoutPanel)
            var flow = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 60,
                BackColor = dominantBg,
                FlowDirection = FlowDirection.LeftToRight,
                Padding = new Padding(20),
                AutoSize = true
            };
            this.Controls.Add(flow);
            flow.BringToFront();
            flow.Controls.Add(btnLoadData);
            flow.Controls.Add(btnExportPDF);
            flow.Controls.Add(btnLoadFromExcel);
        }


        // ------------------------
        // Form Load Event
        // ------------------------
        private void KpiDashboardForm_Load(object sender, EventArgs e)
        {
            this.Text = $"KPI Dashboard - {franchiseName}";

        }

        // ------------------------
        // Behaviors (Methods)
        // ------------------------
        private void LoadKpiData()
       
        {
            dailySales = new List<int> { 75, 90, 85, 40, 100 };

            chartSales.Series[0].Points.Clear();
            string[] days = { "Mon", "Tue", "Wed", "Thu", "Fri" };

            for (int i = 0; i < dailySales.Count; i++)
            {
                chartSales.Series[0].Points.AddXY(days[i], dailySales[i]);
            }
        }
                
        // ------------------------
        // Button Click Events
        // ------------------------
        private void btnRefreshCharts_Click(object sender, EventArgs e)
        {
            {
                if (btnLoadData.Checked)
                {
                    LoadKpiFromAccess();
                }
                else if (btnLoadData.Checked)
                {
                    LoadKpiData();
                }
                else
                {
                    MessageBox.Show("Please select a data source.");
                }
            }
            string dbPath = @"C:\Users\marce\Desktop\Dashboards\FranchiseKPI.accdb";
            string connStr = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};";

            string query = "SELECT Day, Sales, Revenue, Customers FROM tblKPI";

            using (OleDbConnection conn = new OleDbConnection(connStr))
            {
                try
                {
                    conn.Open();
                    OleDbCommand cmd = new OleDbCommand(query, conn);
                    OleDbDataReader reader = cmd.ExecuteReader();

                    chartSales.Series[0].Points.Clear();
                    chartRevenue.Series[0].Points.Clear();
                    chartCustomers.Series[0].Points.Clear();

                    while (reader.Read())
                    {
                        try
                        {
                            string day = reader["Day"].ToString();

                            int sales = reader["Sales"] != DBNull.Value ? Convert.ToInt32(reader["Sales"]) : 0;
                            decimal revenue = reader["Revenue"] != DBNull.Value ? Convert.ToDecimal(reader["Revenue"]) : 0;
                            int customers = reader["Customers"] != DBNull.Value ? Convert.ToInt32(reader["Customers"]) : 0;

                            chartSales.Series[0].Points.AddXY(day, sales);
                            chartRevenue.Series[0].Points.AddXY(day, revenue);
                            chartCustomers.Series[0].Points.AddXY(day, customers);
                        }
                        catch (Exception exInner)
                        {
                            MessageBox.Show("Row error: " + exInner.Message);
                        }
                    }

                    MessageBox.Show("KPI data loaded successfully!");
                }
                catch (Exception exOuter)
                {
                    MessageBox.Show("Database error: " + exOuter.Message);
                }
            }
        }


        private void btnExportPDF_Click(object sender, EventArgs e)
        {
            string sourceLabel = "";

            if (rbAccessData.Checked)
            {
                sourceLabel = "Access";
            }
            else if (rbSampleData.Checked)
            {
                sourceLabel = "Sample";
            }
            else
            {
                MessageBox.Show("Please select a data source.");
                return;
            }

            string fileName = $"KpiReport_{sourceLabel}_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
            string folderPath = Path.Combine(Application.StartupPath, "Reports");

            if (!Directory.Exists(folderPath))
                Directory.CreateDirectory(folderPath);

            string filePath = Path.Combine(folderPath, fileName);

            // Call your existing method to export to PDF (pass filePath)
            ExportChartsToPdf(filePath);

            MessageBox.Show($"Report saved: {filePath}", "Export Successful", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }


        private void ExportChartsToPdf(string filePath)
        {
            PdfDocument document = new PdfDocument();
            document.Info.Title = "Franchise KPI Dashboard";

            // === Page setup ===
            PdfPage page = document.AddPage();
            XGraphics gfx = XGraphics.FromPdfPage(page);

            // === Banner/Header background ===
            XSolidBrush bannerBrush = new XSolidBrush(XColor.FromArgb(255, 255, 255)); // Light background
            gfx.DrawRectangle(bannerBrush, 0, 0, page.Width, 70);

            // === Logo on the left ===
            string logoPath = Path.Combine(Application.StartupPath, "Assets", "Logos", "MES.png"); // Match your app icon
            if (File.Exists(logoPath))
            {
                XImage logo = XImage.FromFile(logoPath);
                gfx.DrawImage(logo, 20, 10, 40, 40); // Small app icon
            }

            // === Dashboard title in center ===
            XFont titleFont = new XFont("Segoe UI", 20, XFontStyleEx.Bold);
            gfx.DrawString("Franchise KPI Dashboard", titleFont, XBrushes.DarkGreen,
                new XRect(0, 20, page.Width, 40), XStringFormats.TopCenter);


            // Fonts
            
            XFont labelFont = new XFont("Arial", 12, XFontStyleEx.Regular);

            // Header & Branding
            string brandPath = Path.Combine(Application.StartupPath, "Assets", "Logos", "Brand.png");
            if (File.Exists(brandPath))
            {
                XImage logo = XImage.FromFile(brandPath);
                gfx.DrawImage(logo, 40, 20, 100, 50);
            }

            gfx.DrawString("Franchise KPI Dashboard", titleFont, XBrushes.DarkSlateGray,
                new XRect(0, 40, page.Width, 50), XStringFormats.TopCenter);

            // Export chart images from controls
            Bitmap bmpSales = new Bitmap(chartSales.Width, chartSales.Height);
            chartSales.DrawToBitmap(bmpSales, new Rectangle(0, 0, bmpSales.Width, bmpSales.Height));

            Bitmap bmpRevenue = new Bitmap(chartRevenue.Width, chartRevenue.Height);
            chartRevenue.DrawToBitmap(bmpRevenue, new Rectangle(0, 0, bmpRevenue.Width, bmpRevenue.Height));

            Bitmap bmpCustomers = new Bitmap(chartCustomers.Width, chartCustomers.Height);
            chartCustomers.DrawToBitmap(bmpCustomers, new Rectangle(0, 0, bmpCustomers.Width, bmpCustomers.Height));

            // Load all logos
            XImage imgSales = XImage.FromFile(Path.Combine(Application.StartupPath, "Assets", "Logos", "Sales.png"));
            XImage imgRevenue = XImage.FromFile(Path.Combine(Application.StartupPath, "Assets", "Logos", "Revenue.png"));
            XImage imgCustomers = XImage.FromFile(Path.Combine(Application.StartupPath, "Assets", "Logos", "Customers.png"));

            // Drawing section layout
            double sectionTop = 100;
            double chartScale = 0.5;

            // Draw sections
            void DrawChartSection(string label, XImage icon, Bitmap chart, ref double top)
            {
                gfx.DrawImage(icon, 50, top - 15, 20, 20);
                gfx.DrawString(label, labelFont, XBrushes.OrangeRed, new XPoint(80, top));
                gfx.DrawImage(chart, 50, top + 10, chart.Width * chartScale, chart.Height * chartScale);
                top += chart.Height * chartScale + 60;
            }

            DrawChartSection("Sales", imgSales, bmpSales, ref sectionTop);
            DrawChartSection("Revenue", imgRevenue, bmpRevenue, ref sectionTop);
            DrawChartSection("Customers", imgCustomers, bmpCustomers, ref sectionTop);


            // === Footer ===
            double footerHeight = 50;
            XSolidBrush footerBrush = new XSolidBrush(XColors.DarkSlateGray);
            gfx.DrawRectangle(footerBrush, 0, page.Height - footerHeight, page.Width, footerHeight);

            // === Footer text (contact info) ===
            string contactText = "© 2025 Country Tech Innovations | support@futureprooftech.biz";
            

            XFont footerFont = new XFont("Segoe UI", 8, XFontStyleEx.Regular);
            gfx.DrawString(contactText, footerFont, XBrushes.Gray, new XRect(0,
                page.Height - 40, page.Width, 20), XStringFormats.Center);

 

            // Save PDF
            document.Save(filePath);

            // Cleanup
            bmpSales.Dispose();
            bmpRevenue.Dispose();
            bmpCustomers.Dispose();
            imgSales.Dispose();
            imgRevenue.Dispose();
            imgCustomers.Dispose();
        }



        // ------------------------
        // Getters and Setters
        // ------------------------
        public string FranchiseName
        {
            get { return franchiseName; }
            set { franchiseName = value; }
        }

        private void btnLoadData_Click(object sender, EventArgs e)
        {
            // Example for chartSales
            chartSales.Series[0].Points.Clear();
            chartSales.Series[0].Points.AddXY("Mon", 75);
            chartSales.Series[0].Points.AddXY("Tue", 90);
            chartSales.Series[0].Points.AddXY("Wed", 85);
            chartSales.Series[0].Points.AddXY("Thu", 40);
            chartSales.Series[0].Points.AddXY("Fri", 100);

            // TODO: chartRevenue.Series[0].Points.AddXY(...) for the 2nd chart
            // TODO: chartCustomerCount.Series[0].Points.AddXY(...) for the 3rd chart
        }

        private void btnLoadFromExcel_Click(object sender, EventArgs e)
        {
            LoadKpiFromAccess();  // this ends here
        }
        // now paste your method here:

        private void LoadKpiFromAccess()
        {

            string dbPath = @"C:\Users\marce\Desktop\Dashboards\FranchiseKPI.accdb";
            string connStr = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};";

            string query = "SELECT Day, Sales, Revenue, Customers FROM tblKPI";

            using (OleDbConnection conn = new OleDbConnection(connStr))
            {
               
                try
                {
                    conn.Open();
                    OleDbCommand cmd = new OleDbCommand(query, conn);
                    OleDbDataReader reader = cmd.ExecuteReader();


                    chartSales.Series[0].Points.Clear();
                    chartRevenue.Series[0].Points.Clear();
                    chartCustomers.Series[0].Points.Clear();

                    while (reader.Read())
                    {
                        string day = reader["Day"].ToString();
                        int sales = Convert.ToInt32(reader["Sales"]);
                        decimal revenue = Convert.ToDecimal(reader["Revenue"]);
                        int customers = Convert.ToInt32(reader["Customers"]);

                        chartSales.Series[0].Points.AddXY(day, sales);
                        chartRevenue.Series[0].Points.AddXY(day, revenue);
                        chartCustomers.Series[0].Points.AddXY(day, customers);
                    }

                    // Repeat for other months or values

                    chartSales.Series[0].ChartType = SeriesChartType.Column;
                    chartRevenue.Series[0].ChartType = SeriesChartType.Column;
                    chartCustomers.Series[0].ChartType = SeriesChartType.Column;


                    while (reader.Read())
                    {
                        string day = reader["Day"].ToString();
                        int sales = Convert.ToInt32(reader["Sales"]);
                        decimal revenue = Convert.ToDecimal(reader["Revenue"]);
                        int customers = Convert.ToInt32(reader["Customers"]);

                        chartSales.Series[0].Points.AddXY(day, sales);
                        chartRevenue.Series[0].Points.AddXY(day, revenue);
                        chartCustomers.Series[0].Points.AddXY(day, customers);
                        chartSales.Series[0].Points.Clear();
                        chartSales.Series[0].Points.AddXY("January", 120);
                        chartSales.Series[0].Points.AddXY("February", 150);
                        // Repeat for other months or values

                        chartSales.Series[0].ChartType = SeriesChartType.Column;
                        chartRevenue.Series[0].ChartType = SeriesChartType.Column;
                        chartCustomers.Series[0].ChartType = SeriesChartType.Column;

                    }



                    MessageBox.Show("KPI data loaded from Access!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message);
                }

            }
        
        }
        private void LoadChartFromDataTable(DataTable dataTable)
        {
            chartSales.Series[0].Points.Clear();
            chartSales.Series[0].ChartType = SeriesChartType.Column;

            foreach (DataRow row in dataTable.Rows)
            {
                chartSales.Series[0].Points.AddXY(row["Month"], row["Sales"]);
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            chartSales.Series[0].ChartType = SeriesChartType.Column;
            chartSales.Series[0].Points.AddXY("Jan", 100);
            chartSales.Series[0].Points.AddXY("Feb", 120);
        }

        private void chartCustomers_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }
    }
}
