using System;
using System.Configuration;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.DirectoryServices.AccountManagement;
using Microsoft.VisualBasic;
using Microsoft.Vbe;
using Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using System.Timers;
using WindowsFormsApplication1;
using System.Data.Sql;
using MySql;
using MySql.Data;
using System.Data.SqlClient;
using System.Drawing;
using MySql.Data.MySqlClient;
using System.IO;
using System.Diagnostics;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        internal Label infolabeltrans;
        internal Label infolabelDis;
        internal Label Infolabelnew;
        internal TableLayoutPanel TableLayoutPanel1;
        internal ComboBox ComboBox1;
        internal Label Serial;
        internal Label Porder;
        internal Label Label10;
        internal Label Label9;
        internal Label Label6;
        internal Label Label7;
        internal Label Worder;
        internal Label Maca;
        internal Label Label11;
        internal Label Label8;
        internal Button CompleteButton;
        internal Button TransButton;
        internal Button DisposalButton;
        internal Button NewButton;
        private ToolStrip toolStrip1;
        private ToolStripDropDownButton toolStripDropDownButton1;
        private ToolStripSeparator toolStripSeparator1;
        private ToolStripLabel DateLabel;
        private ToolStripLabel timelabel;
        private System.Windows.Forms.Timer timer1;
        private IContainer components;
        internal Label assettag;
        internal TextBox TextBox1;
        String ConnectionString;
        MySqlConnection connection;
        MySqlDataAdapter adapter;
        MySqlCommand command;
        private Label labeldept;
        private Label labeldepartment;
        private ToolStripDropDownButton toolStripDropDownButton2;
        private ToolStripDropDownButton toolStripDropDownButton3;
        private ToolStripProgressBar toolStripProgressBar1;
        private ToolStripLabel toolStripLabel1;
        private ToolStripMenuItem newToolStripMenuItem;
        private ToolStripMenuItem closeToolStripMenuItem;
        private ToolStripMenuItem exitToolStripMenuItem;
        private ToolStripMenuItem copyToolStripMenuItem;
        private ToolStripMenuItem pasteToolStripMenuItem;
        private ToolStripMenuItem pushToolStripMenuItem;
        private ToolStripMenuItem pullToolStripMenuItem;
        private ToolStripMenuItem pushToolStripMenuItem1;
        private ToolStripMenuItem pullToolStripMenuItem1;
        private ToolStripMenuItem removeToolStripMenuItem;
        private ListBox listBox1;
        private ToolStripMenuItem daraSheetToolStripMenuItem;
        private ToolStripMenuItem openToolStripMenuItem;
        private ToolStripMenuItem editToolStripMenuItem;
        private ToolStripMenuItem settingsToolStripMenuItem;
        private ToolStripMenuItem searchToolStripMenuItem;
        public Form1()
        {
            InitializeComponent();

            ConnectionString = ConfigurationManager.ConnectionStrings["WindowsFormsApplication1.Properties.Settings.MySQL"].ConnectionString;
                }

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.infolabeltrans = new System.Windows.Forms.Label();
            this.infolabelDis = new System.Windows.Forms.Label();
            this.Infolabelnew = new System.Windows.Forms.Label();
            this.TableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.assettag = new System.Windows.Forms.Label();
            this.ComboBox1 = new System.Windows.Forms.ComboBox();
            this.Serial = new System.Windows.Forms.Label();
            this.Porder = new System.Windows.Forms.Label();
            this.Label10 = new System.Windows.Forms.Label();
            this.Label9 = new System.Windows.Forms.Label();
            this.Label7 = new System.Windows.Forms.Label();
            this.Label11 = new System.Windows.Forms.Label();
            this.Label8 = new System.Windows.Forms.Label();
            this.Maca = new System.Windows.Forms.Label();
            this.labeldept = new System.Windows.Forms.Label();
            this.labeldepartment = new System.Windows.Forms.Label();
            this.Worder = new System.Windows.Forms.Label();
            this.Label6 = new System.Windows.Forms.Label();
            this.CompleteButton = new System.Windows.Forms.Button();
            this.TransButton = new System.Windows.Forms.Button();
            this.DisposalButton = new System.Windows.Forms.Button();
            this.NewButton = new System.Windows.Forms.Button();
            this.TextBox1 = new System.Windows.Forms.TextBox();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.toolStripDropDownButton1 = new System.Windows.Forms.ToolStripDropDownButton();
            this.newToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.closeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripDropDownButton2 = new System.Windows.Forms.ToolStripDropDownButton();
            this.copyToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.pasteToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.pushToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripDropDownButton3 = new System.Windows.Forms.ToolStripDropDownButton();
            this.pullToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.pushToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.pullToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.removeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.searchToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.daraSheetToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.editToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.settingsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripProgressBar1 = new System.Windows.Forms.ToolStripProgressBar();
            this.DateLabel = new System.Windows.Forms.ToolStripLabel();
            this.timelabel = new System.Windows.Forms.ToolStripLabel();
            this.toolStripLabel1 = new System.Windows.Forms.ToolStripLabel();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.TableLayoutPanel1.SuspendLayout();
            this.toolStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // infolabeltrans
            // 
            this.infolabeltrans.AutoSize = true;
            this.infolabeltrans.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.infolabeltrans.Location = new System.Drawing.Point(82, 107);
            this.infolabeltrans.Name = "infolabeltrans";
            this.infolabeltrans.Size = new System.Drawing.Size(208, 32);
            this.infolabeltrans.TabIndex = 38;
            this.infolabeltrans.Text = "Transfer Asset Log";
            this.infolabeltrans.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.infolabeltrans.Visible = false;
            // 
            // infolabelDis
            // 
            this.infolabelDis.AutoSize = true;
            this.infolabelDis.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.infolabelDis.Location = new System.Drawing.Point(80, 107);
            this.infolabelDis.Name = "infolabelDis";
            this.infolabelDis.Size = new System.Drawing.Size(213, 32);
            this.infolabelDis.TabIndex = 37;
            this.infolabelDis.Text = "Disposal Asset Log";
            this.infolabelDis.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.infolabelDis.Visible = false;
            // 
            // Infolabelnew
            // 
            this.Infolabelnew.AutoSize = true;
            this.Infolabelnew.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Infolabelnew.Location = new System.Drawing.Point(104, 107);
            this.Infolabelnew.Name = "Infolabelnew";
            this.Infolabelnew.Size = new System.Drawing.Size(172, 32);
            this.Infolabelnew.TabIndex = 36;
            this.Infolabelnew.Text = "New Asset Log";
            this.Infolabelnew.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.Infolabelnew.Visible = false;
            // 
            // TableLayoutPanel1
            // 
            this.TableLayoutPanel1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.TableLayoutPanel1.BackColor = System.Drawing.SystemColors.Menu;
            this.TableLayoutPanel1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.TableLayoutPanel1.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.Inset;
            this.TableLayoutPanel1.ColumnCount = 2;
            this.TableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.TableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.TableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.TableLayoutPanel1.Controls.Add(this.assettag, 1, 0);
            this.TableLayoutPanel1.Controls.Add(this.ComboBox1, 0, 1);
            this.TableLayoutPanel1.Controls.Add(this.Serial, 1, 3);
            this.TableLayoutPanel1.Controls.Add(this.Porder, 1, 2);
            this.TableLayoutPanel1.Controls.Add(this.Label10, 0, 2);
            this.TableLayoutPanel1.Controls.Add(this.Label9, 0, 3);
            this.TableLayoutPanel1.Controls.Add(this.Label7, 0, 4);
            this.TableLayoutPanel1.Controls.Add(this.Label11, 0, 5);
            this.TableLayoutPanel1.Controls.Add(this.Label8, 0, 0);
            this.TableLayoutPanel1.Controls.Add(this.Maca, 1, 5);
            this.TableLayoutPanel1.Controls.Add(this.labeldept, 0, 6);
            this.TableLayoutPanel1.Controls.Add(this.labeldepartment, 1, 6);
            this.TableLayoutPanel1.Controls.Add(this.Worder, 1, 4);
            this.TableLayoutPanel1.Controls.Add(this.Label6, 1, 1);
            this.TableLayoutPanel1.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TableLayoutPanel1.Location = new System.Drawing.Point(35, 250);
            this.TableLayoutPanel1.MinimumSize = new System.Drawing.Size(100, 300);
            this.TableLayoutPanel1.Name = "TableLayoutPanel1";
            this.TableLayoutPanel1.Padding = new System.Windows.Forms.Padding(8);
            this.TableLayoutPanel1.RowCount = 7;
            this.TableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 14.28571F));
            this.TableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 14.28571F));
            this.TableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 14.28571F));
            this.TableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 14.28571F));
            this.TableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 14.28571F));
            this.TableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 14.28571F));
            this.TableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 14.28571F));
            this.TableLayoutPanel1.Size = new System.Drawing.Size(400, 514);
            this.TableLayoutPanel1.TabIndex = 35;
            // 
            // assettag
            // 
            this.assettag.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.assettag.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.assettag.Location = new System.Drawing.Point(208, 12);
            this.assettag.Margin = new System.Windows.Forms.Padding(0);
            this.assettag.Name = "assettag";
            this.assettag.Size = new System.Drawing.Size(175, 63);
            this.assettag.TabIndex = 4;
            this.assettag.Text = "Label4";
            this.assettag.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.assettag.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.assettag_MouseDoubleClick);
            // 
            // ComboBox1
            // 
            this.ComboBox1.AllowDrop = true;
            this.ComboBox1.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.ComboBox1.DropDownWidth = 150;
            this.ComboBox1.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.ComboBox1.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ComboBox1.FormattingEnabled = true;
            this.ComboBox1.Items.AddRange(new object[] {
            "Intel NUC 6i3SYK",
            "Dell Latitude E7450",
            "Symbol Tech STB4278 Scanner",
            "Dell Poweredge R620 Server",
            "Dell 23\" Monitor",
            "Cisco AIR-LAP 1142N-A-K9 Access Point"});
            this.ComboBox1.Location = new System.Drawing.Point(41, 94);
            this.ComboBox1.Margin = new System.Windows.Forms.Padding(7, 6, 7, 6);
            this.ComboBox1.Name = "ComboBox1";
            this.ComboBox1.Size = new System.Drawing.Size(126, 40);
            this.ComboBox1.TabIndex = 24;
            this.ComboBox1.Text = "Description";
            this.ComboBox1.SelectedIndexChanged += new System.EventHandler(this.ComboBox1_SelectedIndexChanged);
            this.ComboBox1.Click += new System.EventHandler(this.ComboBox1_Click);
            this.ComboBox1.Enter += new System.EventHandler(this.ComboBox1_Enter);
            this.ComboBox1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ComboBox1_KeyDown);
            this.ComboBox1.Leave += new System.EventHandler(this.ComboBox1_Leave);
            // 
            // Serial
            // 
            this.Serial.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.Serial.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Serial.Location = new System.Drawing.Point(208, 225);
            this.Serial.Margin = new System.Windows.Forms.Padding(0);
            this.Serial.Name = "Serial";
            this.Serial.Size = new System.Drawing.Size(175, 57);
            this.Serial.TabIndex = 5;
            this.Serial.Text = "Label3";
            this.Serial.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.Serial.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.Serial_MouseDoubleClick);
            // 
            // Porder
            // 
            this.Porder.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.Porder.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Porder.Location = new System.Drawing.Point(208, 153);
            this.Porder.Margin = new System.Windows.Forms.Padding(0);
            this.Porder.Name = "Porder";
            this.Porder.Size = new System.Drawing.Size(175, 61);
            this.Porder.TabIndex = 6;
            this.Porder.Text = "Label3";
            this.Porder.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.Porder.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.Porder_MouseDoubleClick);
            // 
            // Label10
            // 
            this.Label10.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.Label10.AutoSize = true;
            this.Label10.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Label10.Location = new System.Drawing.Point(60, 168);
            this.Label10.Margin = new System.Windows.Forms.Padding(7, 0, 7, 0);
            this.Label10.Name = "Label10";
            this.Label10.Size = new System.Drawing.Size(89, 32);
            this.Label10.TabIndex = 20;
            this.Label10.Text = "Purch#";
            // 
            // Label9
            // 
            this.Label9.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.Label9.AutoSize = true;
            this.Label9.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Label9.Location = new System.Drawing.Point(61, 238);
            this.Label9.Margin = new System.Windows.Forms.Padding(7, 0, 7, 0);
            this.Label9.Name = "Label9";
            this.Label9.Size = new System.Drawing.Size(87, 32);
            this.Label9.TabIndex = 19;
            this.Label9.Text = "Serial#";
            // 
            // Label7
            // 
            this.Label7.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.Label7.AutoSize = true;
            this.Label7.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Label7.Location = new System.Drawing.Point(70, 308);
            this.Label7.Margin = new System.Windows.Forms.Padding(7, 0, 7, 0);
            this.Label7.Name = "Label7";
            this.Label7.Size = new System.Drawing.Size(69, 32);
            this.Label7.TabIndex = 17;
            this.Label7.Text = "WO#";
            // 
            // Label11
            // 
            this.Label11.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.Label11.AutoSize = true;
            this.Label11.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Label11.Location = new System.Drawing.Point(64, 378);
            this.Label11.Margin = new System.Windows.Forms.Padding(7, 0, 7, 0);
            this.Label11.Name = "Label11";
            this.Label11.Size = new System.Drawing.Size(81, 32);
            this.Label11.TabIndex = 21;
            this.Label11.Text = "MAC#";
            // 
            // Label8
            // 
            this.Label8.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.Label8.AutoSize = true;
            this.Label8.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Label8.Location = new System.Drawing.Point(62, 28);
            this.Label8.Margin = new System.Windows.Forms.Padding(7, 0, 7, 0);
            this.Label8.Name = "Label8";
            this.Label8.Size = new System.Drawing.Size(85, 32);
            this.Label8.TabIndex = 18;
            this.Label8.Text = "Asset#";
            this.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // Maca
            // 
            this.Maca.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.Maca.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Maca.Location = new System.Drawing.Point(208, 360);
            this.Maca.Margin = new System.Windows.Forms.Padding(0);
            this.Maca.Name = "Maca";
            this.Maca.Size = new System.Drawing.Size(175, 68);
            this.Maca.TabIndex = 7;
            this.Maca.Text = "Label3";
            this.Maca.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.Maca.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.Maca_MouseDoubleClick);
            // 
            // labeldept
            // 
            this.labeldept.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.labeldept.AutoSize = true;
            this.labeldept.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labeldept.Location = new System.Drawing.Point(33, 451);
            this.labeldept.Name = "labeldept";
            this.labeldept.Size = new System.Drawing.Size(143, 32);
            this.labeldept.TabIndex = 25;
            this.labeldept.Text = "Department";
            // 
            // labeldepartment
            // 
            this.labeldepartment.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.labeldepartment.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labeldepartment.Location = new System.Drawing.Point(204, 430);
            this.labeldepartment.Margin = new System.Windows.Forms.Padding(0);
            this.labeldepartment.Name = "labeldepartment";
            this.labeldepartment.Size = new System.Drawing.Size(183, 74);
            this.labeldepartment.TabIndex = 26;
            this.labeldepartment.Text = "label1";
            this.labeldepartment.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.labeldepartment.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.labeldepartment_MouseDoubleClick);
            // 
            // Worder
            // 
            this.Worder.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.Worder.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Worder.Location = new System.Drawing.Point(208, 296);
            this.Worder.Margin = new System.Windows.Forms.Padding(0);
            this.Worder.Name = "Worder";
            this.Worder.Size = new System.Drawing.Size(175, 56);
            this.Worder.TabIndex = 3;
            this.Worder.Text = "wo";
            this.Worder.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.Worder.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.Worder_MouseDoubleClick);
            // 
            // Label6
            // 
            this.Label6.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.Label6.BackColor = System.Drawing.Color.PaleGreen;
            this.Label6.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Label6.Location = new System.Drawing.Point(201, 85);
            this.Label6.Margin = new System.Windows.Forms.Padding(0);
            this.Label6.Name = "Label6";
            this.Label6.Size = new System.Drawing.Size(189, 57);
            this.Label6.TabIndex = 16;
            this.Label6.Text = "Label6";
            this.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.Label6.Visible = false;
            this.Label6.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.Label6_MouseDoubleClick);
            this.Label6.MouseHover += new System.EventHandler(this.Label6_MouseHover);
            // 
            // CompleteButton
            // 
            this.CompleteButton.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.CompleteButton.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.CompleteButton.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CompleteButton.Location = new System.Drawing.Point(136, 193);
            this.CompleteButton.Margin = new System.Windows.Forms.Padding(4);
            this.CompleteButton.Name = "CompleteButton";
            this.CompleteButton.Size = new System.Drawing.Size(80, 35);
            this.CompleteButton.TabIndex = 34;
            this.CompleteButton.Text = "Complete";
            this.CompleteButton.UseVisualStyleBackColor = false;
            this.CompleteButton.Click += new System.EventHandler(this.CompleteButton_Click);
            // 
            // TransButton
            // 
            this.TransButton.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.TransButton.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.TransButton.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TransButton.Location = new System.Drawing.Point(282, 53);
            this.TransButton.Margin = new System.Windows.Forms.Padding(4);
            this.TransButton.Name = "TransButton";
            this.TransButton.Size = new System.Drawing.Size(80, 35);
            this.TransButton.TabIndex = 33;
            this.TransButton.Text = "Transfer";
            this.TransButton.UseVisualStyleBackColor = false;
            this.TransButton.Click += new System.EventHandler(this.TransButton_Click);
            // 
            // DisposalButton
            // 
            this.DisposalButton.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.DisposalButton.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.DisposalButton.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.DisposalButton.Location = new System.Drawing.Point(144, 53);
            this.DisposalButton.Margin = new System.Windows.Forms.Padding(4);
            this.DisposalButton.Name = "DisposalButton";
            this.DisposalButton.Size = new System.Drawing.Size(80, 35);
            this.DisposalButton.TabIndex = 32;
            this.DisposalButton.Text = "Disposal";
            this.DisposalButton.UseVisualStyleBackColor = false;
            this.DisposalButton.Click += new System.EventHandler(this.DisposalButton_Click);
            // 
            // NewButton
            // 
            this.NewButton.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.NewButton.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.NewButton.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.NewButton.Location = new System.Drawing.Point(23, 58);
            this.NewButton.Margin = new System.Windows.Forms.Padding(4);
            this.NewButton.Name = "NewButton";
            this.NewButton.Size = new System.Drawing.Size(80, 35);
            this.NewButton.TabIndex = 31;
            this.NewButton.Text = "New";
            this.NewButton.UseVisualStyleBackColor = false;
            this.NewButton.Click += new System.EventHandler(this.NewButton_Click);
            // 
            // TextBox1
            // 
            this.TextBox1.Location = new System.Drawing.Point(77, 152);
            this.TextBox1.Margin = new System.Windows.Forms.Padding(8);
            this.TextBox1.Name = "TextBox1";
            this.TextBox1.Size = new System.Drawing.Size(251, 39);
            this.TextBox1.TabIndex = 30;
            this.TextBox1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textbox1_keydown);
            // 
            // toolStrip1
            // 
            this.toolStrip1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.toolStrip1.ImageScalingSize = new System.Drawing.Size(32, 32);
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripDropDownButton1,
            this.toolStripDropDownButton2,
            this.toolStripDropDownButton3,
            this.toolStripSeparator1,
            this.toolStripProgressBar1,
            this.DateLabel,
            this.timelabel,
            this.toolStripLabel1});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(1010, 47);
            this.toolStrip1.TabIndex = 39;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // toolStripDropDownButton1
            // 
            this.toolStripDropDownButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripDropDownButton1.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.newToolStripMenuItem,
            this.closeToolStripMenuItem,
            this.exitToolStripMenuItem});
            this.toolStripDropDownButton1.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.toolStripDropDownButton1.Image = ((System.Drawing.Image)(resources.GetObject("toolStripDropDownButton1.Image")));
            this.toolStripDropDownButton1.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripDropDownButton1.Name = "toolStripDropDownButton1";
            this.toolStripDropDownButton1.Size = new System.Drawing.Size(74, 44);
            this.toolStripDropDownButton1.Text = "File";
            // 
            // newToolStripMenuItem
            // 
            this.newToolStripMenuItem.Name = "newToolStripMenuItem";
            this.newToolStripMenuItem.Size = new System.Drawing.Size(172, 38);
            this.newToolStripMenuItem.Text = "New";
            // 
            // closeToolStripMenuItem
            // 
            this.closeToolStripMenuItem.Name = "closeToolStripMenuItem";
            this.closeToolStripMenuItem.Size = new System.Drawing.Size(172, 38);
            this.closeToolStripMenuItem.Text = "Close";
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(172, 38);
            this.exitToolStripMenuItem.Text = "Exit";
            // 
            // toolStripDropDownButton2
            // 
            this.toolStripDropDownButton2.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripDropDownButton2.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.copyToolStripMenuItem,
            this.pasteToolStripMenuItem,
            this.pushToolStripMenuItem});
            this.toolStripDropDownButton2.Image = ((System.Drawing.Image)(resources.GetObject("toolStripDropDownButton2.Image")));
            this.toolStripDropDownButton2.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripDropDownButton2.Name = "toolStripDropDownButton2";
            this.toolStripDropDownButton2.Size = new System.Drawing.Size(77, 44);
            this.toolStripDropDownButton2.Text = "Edit";
            this.toolStripDropDownButton2.Click += new System.EventHandler(this.toolStripDropDownButton2_Click);
            // 
            // copyToolStripMenuItem
            // 
            this.copyToolStripMenuItem.Name = "copyToolStripMenuItem";
            this.copyToolStripMenuItem.Size = new System.Drawing.Size(169, 38);
            this.copyToolStripMenuItem.Text = "Copy";
            // 
            // pasteToolStripMenuItem
            // 
            this.pasteToolStripMenuItem.Name = "pasteToolStripMenuItem";
            this.pasteToolStripMenuItem.Size = new System.Drawing.Size(169, 38);
            this.pasteToolStripMenuItem.Text = "Paste";
            // 
            // pushToolStripMenuItem
            // 
            this.pushToolStripMenuItem.Name = "pushToolStripMenuItem";
            this.pushToolStripMenuItem.Size = new System.Drawing.Size(169, 38);
            this.pushToolStripMenuItem.Text = "Push";
            // 
            // toolStripDropDownButton3
            // 
            this.toolStripDropDownButton3.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripDropDownButton3.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.pullToolStripMenuItem,
            this.searchToolStripMenuItem,
            this.daraSheetToolStripMenuItem});
            this.toolStripDropDownButton3.Image = ((System.Drawing.Image)(resources.GetObject("toolStripDropDownButton3.Image")));
            this.toolStripDropDownButton3.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripDropDownButton3.Name = "toolStripDropDownButton3";
            this.toolStripDropDownButton3.Size = new System.Drawing.Size(88, 44);
            this.toolStripDropDownButton3.Text = "View";
            // 
            // pullToolStripMenuItem
            // 
            this.pullToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.pushToolStripMenuItem1,
            this.pullToolStripMenuItem1,
            this.removeToolStripMenuItem});
            this.pullToolStripMenuItem.Name = "pullToolStripMenuItem";
            this.pullToolStripMenuItem.Size = new System.Drawing.Size(324, 38);
            this.pullToolStripMenuItem.Text = "Admin";
            this.pullToolStripMenuItem.Click += new System.EventHandler(this.pullToolStripMenuItem_Click);
            // 
            // pushToolStripMenuItem1
            // 
            this.pushToolStripMenuItem1.Name = "pushToolStripMenuItem1";
            this.pushToolStripMenuItem1.Size = new System.Drawing.Size(200, 38);
            this.pushToolStripMenuItem1.Text = "Push";
            // 
            // pullToolStripMenuItem1
            // 
            this.pullToolStripMenuItem1.Name = "pullToolStripMenuItem1";
            this.pullToolStripMenuItem1.Size = new System.Drawing.Size(200, 38);
            this.pullToolStripMenuItem1.Text = "Pull";
            // 
            // removeToolStripMenuItem
            // 
            this.removeToolStripMenuItem.Name = "removeToolStripMenuItem";
            this.removeToolStripMenuItem.Size = new System.Drawing.Size(200, 38);
            this.removeToolStripMenuItem.Text = "Remove";
            // 
            // searchToolStripMenuItem
            // 
            this.searchToolStripMenuItem.Name = "searchToolStripMenuItem";
            this.searchToolStripMenuItem.Size = new System.Drawing.Size(324, 38);
            this.searchToolStripMenuItem.Text = "Search";
            // 
            // daraSheetToolStripMenuItem
            // 
            this.daraSheetToolStripMenuItem.DoubleClickEnabled = true;
            this.daraSheetToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openToolStripMenuItem,
            this.editToolStripMenuItem,
            this.settingsToolStripMenuItem});
            this.daraSheetToolStripMenuItem.Name = "daraSheetToolStripMenuItem";
            this.daraSheetToolStripMenuItem.Size = new System.Drawing.Size(324, 38);
            this.daraSheetToolStripMenuItem.Text = "Data Sheet";
            // 
            // openToolStripMenuItem
            // 
            this.openToolStripMenuItem.Name = "openToolStripMenuItem";
            this.openToolStripMenuItem.Size = new System.Drawing.Size(324, 38);
            this.openToolStripMenuItem.Text = "Open";
            this.openToolStripMenuItem.Click += new System.EventHandler(this.openToolStripMenuItem_Click);
            // 
            // editToolStripMenuItem
            // 
            this.editToolStripMenuItem.Name = "editToolStripMenuItem";
            this.editToolStripMenuItem.Size = new System.Drawing.Size(324, 38);
            this.editToolStripMenuItem.Text = "Edit";
            this.editToolStripMenuItem.Click += new System.EventHandler(this.editToolStripMenuItem_Click);
            // 
            // settingsToolStripMenuItem
            // 
            this.settingsToolStripMenuItem.Name = "settingsToolStripMenuItem";
            this.settingsToolStripMenuItem.Size = new System.Drawing.Size(324, 38);
            this.settingsToolStripMenuItem.Text = "Settings";
            this.settingsToolStripMenuItem.Click += new System.EventHandler(this.settingsToolStripMenuItem_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 47);
            // 
            // toolStripProgressBar1
            // 
            this.toolStripProgressBar1.Name = "toolStripProgressBar1";
            this.toolStripProgressBar1.Size = new System.Drawing.Size(100, 44);
            // 
            // DateLabel
            // 
            this.DateLabel.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.DateLabel.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.DateLabel.Name = "DateLabel";
            this.DateLabel.Size = new System.Drawing.Size(174, 44);
            this.DateLabel.Text = "toolStripLabel1";
            // 
            // timelabel
            // 
            this.timelabel.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.timelabel.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.timelabel.Name = "timelabel";
            this.timelabel.Size = new System.Drawing.Size(174, 44);
            this.timelabel.Text = "toolStripLabel2";
            // 
            // toolStripLabel1
            // 
            this.toolStripLabel1.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.toolStripLabel1.Name = "toolStripLabel1";
            this.toolStripLabel1.Size = new System.Drawing.Size(174, 44);
            this.toolStripLabel1.Text = "toolStripLabel1";
            // 
            // timer1
            // 
            this.timer1.Enabled = true;
            this.timer1.Interval = 1000;
            this.timer1.Tick += new System.EventHandler(this.timer1_tick);
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.ItemHeight = 32;
            this.listBox1.Location = new System.Drawing.Point(608, 107);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(368, 260);
            this.listBox1.TabIndex = 40;
            // 
            // Form1
            // 
            this.BackColor = System.Drawing.SystemColors.Control;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.ClientSize = new System.Drawing.Size(1010, 816);
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.toolStrip1);
            this.Controls.Add(this.Infolabelnew);
            this.Controls.Add(this.TableLayoutPanel1);
            this.Controls.Add(this.CompleteButton);
            this.Controls.Add(this.TransButton);
            this.Controls.Add(this.DisposalButton);
            this.Controls.Add(this.NewButton);
            this.Controls.Add(this.TextBox1);
            this.Controls.Add(this.infolabeltrans);
            this.Controls.Add(this.infolabelDis);
            this.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "Asset Tracker";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.TableLayoutPanel1.ResumeLayout(false);
            this.TableLayoutPanel1.PerformLayout();
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        Asset AssetNew0 = new Asset();
        Int32 type = 0;
        string assettype = "Null";
        public Int32 currentWHasset = 81039;
        public Int32 currentBCHasset = 94455;
        public Int32 currentWO = 488157;
        public Int32 currentPO = 223591;
        string WO = "0";
        string PO = "0";
        string SN = "0";
        string MAC = "0";
        string Asset = "0";
        string department = "0";
        public int facility = 0;
        string transferfrom = "0";
        string transferto = "0";
        string dateran = "0";
        string datecomplete = "0";
        string[] devices;
        bool passing = false;
        string description = "null";
        int totalitems = 0;
        string scannedtemp = "";
        Boolean authenticated = false;
        Boolean authorized = false;
        string user;

        SqlDataAdapter DBAsset= new SqlDataAdapter();



        private void grabbarcode()
        {
            scannedtemp = TextBox1.Text;
            scannedtemp = scannedtemp.ToLower();
            Grab NewG = new Grab(scannedtemp);

            if (NewG.location == "asset")
            {
                AssetNew0.Tag = NewG.localint;
                assettag.Text = scannedtemp;
            }
            else if (NewG.location == "po")
            {
                AssetNew0.Po = NewG.localint;
                Porder.Text = scannedtemp;

            }
            else if (NewG.location == "sn")
            {
                AssetNew0.Sn = NewG.localstring;
                Serial.Text = scannedtemp;

            }
            else if (NewG.location == "mac")
            {
                AssetNew0.Mac = NewG.localstring;
                Maca.Text = scannedtemp;

            }
            else if (NewG.location == "dept"){
                AssetNew0.Dept = NewG.localstring;
                labeldepartment.Text = scannedtemp;
            }
            else if (NewG.location == "wo")
            {
                AssetNew0.Wo = NewG.localint;
                Worder.Text = scannedtemp;
            }
            else if (NewG.location =="null")
            {

            }
            else
            {


            }
          
           

        }//Used to grab to scan after the enter key is hit

      

        private void textbox1_keydown(object sender, System.Windows.Forms.KeyEventArgs b)
        {
            if (((b.KeyCode == Keys.Enter) || (b.KeyCode == Keys.Return)))
            //Checks for Enter key pressed or Return key pressed
            {
                grabbarcode();//Calls method 
                TextBox1.Clear();//Clears text box input
                TextBox1.Focus();//Focuses on text box
            }

        }

        private void descriptionselected(object sender, EventArgs e)
        {
            ComboBox1.Text = "Description";//Sets text to description
        }//Used to set defualt text

        private void Form1_Load(object sender, EventArgs e)//Form Load
        {
           //string user=UserPrincipal.Current.DisplayName;
            toolStripLabel1.Text = user;
            TextBox1.Visible = false;//Hides the text box
            CompleteButton.Visible = false;//Hides the complete button
            dateran = DateTime.Now.ToString("dddd dd MMMM, yyyy");//Grabs the date ran
            ComboBox1.Visible = false;
            DateLabel.Text = dateran;//Sets the date ran label
            TextBox1.Focus();//Focuses on teh textbox
            timer1_tick(null, null);//Calls the timer to start
           // pulldata();
            
          


        }

        private void NewButton_Click(object sender, EventArgs e)//Called if NEW ASSET button selected 
        {
            //Makes new asset form inputs visible and hides the rest
            Infolabelnew.Visible = true;
            infolabelDis.Visible = false;
            infolabeltrans.Visible = false;
            ComboBox1.Visible = true;
            type = 1;
            assettype = "New";
            NewButton.Visible = false;
            DisposalButton.Visible = false;
            TransButton.Visible = false;
            TextBox1.Visible = true;
            CompleteButton.Visible = true;
            TextBox1.Focus();
            
            
        }

        private void DisposalButton_Click(object sender, EventArgs e)//Called if DISPOSAL button selected
        {
            //Makes disposal asset form inputs visible and hides the rest
            infolabelDis.Visible = true;
            Infolabelnew.Visible = false;
            ComboBox1.Visible = true;
            infolabeltrans.Visible = false;
            type = 2;
            assettype = "Disposal";
            NewButton.Visible = false;
            DisposalButton.Visible = false;
            TransButton.Visible = false;
            TextBox1.Visible = true;
            CompleteButton.Visible = true;
            TextBox1.Focus();
        }

        private void TransButton_Click(object sender, EventArgs e)//Called if TRANSFER button selected
        {
            //Makes transfer asset form inputs visible and hides the rest
            infolabeltrans.Visible = true;
            infolabelDis.Visible = false;
            Infolabelnew.Visible = false;
            ComboBox1.Visible = true;
            type = 3;
            assettype = "Transfer";
            NewButton.Visible = false;
            DisposalButton.Visible = false;
            TransButton.Visible = false;
            TextBox1.Visible = true;
            CompleteButton.Visible = true;
            TextBox1.Focus();
            Label10.Text = "Asset to";
            Label8.Text = "Asset from";
        }

        private void CompleteButton_Click(object sender, EventArgs e)//Called once complete button is pressed 
        {
            bool cleared = true;//Bool to capture if everything is scanned
            string warningmessage = "";
            if ((type == 1 ||type==2))//Checks if New button was selected
            {
                if ((Asset == "0"))//Verifies the Asset was scanned
                {
                    if ((warningmessage == ""))//Checks if warning message is blank
                    {
                        warningmessage = "Please scan the Asset Tag";//If blank it starts the message
                    }
                    else {
                        warningmessage = (warningmessage + ", Asset Tag");//If not blank if adds to the message
                        cleared = false;//Set the Bool to not clear
                    }

                }

                if ((SN == "0"))//Checks if Serial number is blank
                {
                    if ((warningmessage == ""))//Checks if warning message is blank
                    {
                        warningmessage = "Please scan the Serial Number";//Created new messgae
                    }
                    else {
                        warningmessage = (warningmessage + ", Serial Number");//Adds to message
                        cleared = false;//Sets Bool to false
                    }

                }

                if ((PO == "0"&&type!=2))//Checks if PO is scanned
                {
                    if ((warningmessage == ""))//Checks if warning message is blank
                    {
                        warningmessage = "Please scan the PO Number";//Sets the warning message
                    }
                    else {
                        warningmessage = (warningmessage + ", PO Number");//Adds to the warning message
                        cleared = false;//Sets Bool to false
                    }

                    cleared = false;//Sets cleared to false
                
                }

                if ((description == "null"))//checks for blank description
                {
                    if ((warningmessage == ""))//Checks for blank warning message
                    {
                        warningmessage = "Please enter the Item Description";//Creates new message
                    }
                    else {
                        warningmessage = (warningmessage + ", Item Description");//Adds to message
                        cleared = false;//Sets cleared to false
                    }

                    cleared = false;//Sets cleared to false
                }
                toolStripProgressBar1.Value = 15;

            }
            else if ((type == 3))
            {
                if ((Asset == "0"))
                {
                    if ((warningmessage == ""))
                    {
                        warningmessage = "Please scan the Asset From Tag";
                    }
                    else {
                        warningmessage = (warningmessage + ", Asset Tag");
                        cleared = false;
                    }

                    cleared = false;
                }

                if ((PO == "0"))
                {
                    if ((warningmessage == ""))
                    {
                        warningmessage = "Please scan the Asset To Tag";
                    }
                    else {
                        warningmessage = (warningmessage + ", Asset Tag");
                        cleared = false;
                    }

                    cleared = false;
                }

                if ((SN == "0"))
                {
                    if ((warningmessage == ""))
                    {
                        warningmessage = "Please scan the Serial Number";
                    }
                    else {
                        warningmessage = (warningmessage + ", Serial Number");
                        cleared = false;
                    }

                }

                if ((description == "null"))
                {
                    if ((warningmessage == ""))
                    {
                        warningmessage = "Please enter the Item Description";
                    }
                    else {
                        warningmessage = (warningmessage + ", Item Description");
                        cleared = false;
                    }

                    cleared = false;
                }
                toolStripProgressBar1.Value = 15;

            }
            else {

            }

            if ((warningmessage == "Please scan the PO Number"))
            {
                object manualinput = MessageBox.Show("Please scan or type the PO number or other information", "PO");
                object correctanswer = MessageBox.Show(("Is "
                                + (manualinput + " correct?")), "PO", MessageBoxButtons.YesNo);
                if ((correctanswer.ToString() == "6"))
                {
                    PO = manualinput.ToString();
                    Porder.Text = manualinput.ToString();
                    cleared = true;
                }
                else if ((correctanswer.ToString() == "no"))
                {
                    MessageBox.Show(warningmessage);
                }

            }
            else {

            }
            
            if (((cleared == true)
                        && (type == 1)))
            {
               

                string answer = MessageBox.Show("Do you have any further items on the same PO", "Test", MessageBoxButtons.YesNo).ToString();
                //^^Asks user if they have any further devices on this WO & PO
                if ((answer == "no" || answer=="No"))//If no it clears all labels and sends the data to be formated
                {
                    Worder.Text = "";
                    assettag.Text = "";
                    Serial.Text = "";
                    Porder.Text = "";
                    Maca.Text = "";
                    ComboBox1.ResetText();
                    ComboBox1.Text = "Descriptiopn";
                    Label6.Text = "";
                    TextBox1.Focus();
                    dategrab();
                    formatmessage();
                    wordformater();
                    System.Diagnostics.Process.Start("E:\\C#\\WriteText.txt");

                    toolStripProgressBar1.Value = 30;

                }
                else if ((answer == "yes" || answer=="Yes" ))//If yes clears some labels and waits to send data for Processing 
                {
                    wordformater();
                    formatmessage();
                 //   if (Label6.Text=) Needs to add ticker to 3d Array for correct message formating
                    totalitems = totalitems + 1;
                    assettag.Text = "";
                    Asset = "0";
                    Serial.Text = "";
                    SN = "0";
                    Maca.Text = "";
                    MAC = "0";
                    this.Focus();
                    TextBox1.Focus();
                }

            }
            else if ((cleared == false))//Checks if any errors occured
            {
                MessageBox.Show(warningmessage);//Displays messgae
            }

        }

        private void wordformater()//Called to format the word doc form
        {
            Microsoft.Office.Interop.Word.Application objWordApp;
            objWordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document objDoc;
            if (type == 1)
            {
                objWordApp.Documents.Open("E:\\c#\\New.docx");
            }
            else if(type == 2){
                objWordApp.Documents.Open("E:\\c#\\Disposal.docx");
            }
            else
            {
                objWordApp.Documents.Open("E:\\c#\\Transfer.docx");
            }
            objDoc = objWordApp.ActiveDocument;
            // Find and replace some text.
            object MatchCase = true;
            object MatchWholeWord = true;
            object  MatchWildCards= false;
            object MatchSoundsLike  = false;
            object  MatchAllWordForms= false;
            object Forward = true;
            object Format = false;
            object MatchKashida = false;
            object MatchDiacritics = false;
            object  MatchAlefHamza= false;
            object  MatchControl= false;
            object  Read_Only= false;
            object  Visible = false;
            object  Replace = 2;
            object Wrap = 1;

           


           object Find = "<DATE>";
            object ReplaceWith = dateran;
            objDoc.Content.Find.Execute(ref Find,ref MatchCase,ref MatchWholeWord,ref MatchWildCards,ref MatchSoundsLike, ref MatchAllWordForms, ref Forward, ref Wrap,ref Format, ref ReplaceWith,ref Replace, ref MatchKashida, ref MatchDiacritics, ref MatchAlefHamza, ref MatchControl);

            // TODO: Labeled Arguments not supported. Argument: 1 := 'FindText'
            // TODO: Labeled Arguments not supported. Argument: 2 := 'ReplaceWith'
            // TODO: Labeled Arguments not supported. Argument: 3 := 'Replace'

            Find = "<NAME>";
            ReplaceWith = user;

            objDoc.Content.Find.Execute(ref Find, ref MatchCase, ref MatchWholeWord, ref MatchWildCards, ref MatchSoundsLike, ref MatchAllWordForms, ref Forward, ref Wrap, ref Format, ref ReplaceWith, ref Replace, ref MatchKashida, ref MatchDiacritics, ref MatchAlefHamza, ref MatchControl);
            // TODO: Labeled Arguments not supported. Argument: 1 := 'FindText'
            // TODO: Labeled Arguments not supported. Argument: 2 := 'ReplaceWith'
            // TODO: Labeled Arguments not supported. Argument: 3 := 'Replace'

            Find = "<DATERAN>";
            ReplaceWith = datecomplete;
            objDoc.Content.Find.Execute(ref Find, ref MatchCase, ref MatchWholeWord, ref MatchWildCards, ref MatchSoundsLike, ref MatchAllWordForms, ref Forward, ref Wrap, ref Format, ref ReplaceWith, ref Replace, ref MatchKashida, ref MatchDiacritics, ref MatchAlefHamza, ref MatchControl);
            // TODO: Labeled Arguments not supported. Argument: 1 := 'FindText'
            // TODO: Labeled Arguments not supported. Argument: 2 := 'ReplaceWith'
            // TODO: Labeled Arguments not supported. Argument: 3 := 'Replace'
            if (department != "0")
            {
                Find = "<DEPT>";
                ReplaceWith = department;
                objDoc.Content.Find.Execute(ref Find, ref MatchCase, ref MatchWholeWord, ref MatchWildCards, ref MatchSoundsLike, ref MatchAllWordForms, ref Forward, ref Wrap, ref Format, ref ReplaceWith, ref Replace, ref MatchKashida, ref MatchDiacritics, ref MatchAlefHamza, ref MatchControl);
                // TODO: Labeled Arguments not supported. Argument: 1 := 'FindText'
                // TODO: Labeled Arguments not supported. Argument: 2 := 'ReplaceWith'
                // TODO: Labeled Arguments not supported. Argument: 3 := 'Replace'
            }
            if ((type == 1))
            {
                Find = "N;";
                ReplaceWith = "/";
                objDoc.Content.Find.Execute(ref Find, ref MatchCase, ref MatchWholeWord, ref MatchWildCards, ref MatchSoundsLike, ref MatchAllWordForms, ref Forward, ref Wrap, ref Format, ref ReplaceWith, ref Replace, ref MatchKashida, ref MatchDiacritics, ref MatchAlefHamza, ref MatchControl);
                // TODO: Labeled Arguments not supported. Argument: 1 := 'FindText'
                // TODO: Labeled Arguments not supported. Argument: 2 := 'ReplaceWith'
                // TODO: Labeled Arguments not supported. Argument: 3 := 'Replace'

                Find = "D;";
                ReplaceWith = "";
                objDoc.Content.Find.Execute(ref Find, ref MatchCase, ref MatchWholeWord, ref MatchWildCards, ref MatchSoundsLike, ref MatchAllWordForms, ref Forward, ref Wrap, ref Format, ref ReplaceWith, ref Replace, ref MatchKashida, ref MatchDiacritics, ref MatchAlefHamza, ref MatchControl);
                // TODO: Labeled Arguments not supported. Argument: 1 := 'FindText'
                // TODO: Labeled Arguments not supported. Argument: 2 := 'ReplaceWith'
                // TODO: Labeled Arguments not supported. Argument: 3 := 'Replace'

                Find = "T;";
                ReplaceWith = "";
                objDoc.Content.Find.Execute(ref Find, ref MatchCase, ref MatchWholeWord, ref MatchWildCards, ref MatchSoundsLike, ref MatchAllWordForms, ref Forward, ref Wrap, ref Format, ref ReplaceWith, ref Replace, ref MatchKashida, ref MatchDiacritics, ref MatchAlefHamza, ref MatchControl);
                // TODO: Labeled Arguments not supported. Argument: 1 := 'FindText'
                // TODO: Labeled Arguments not supported. Argument: 2 := 'ReplaceWith'
                // TODO: Labeled Arguments not supported. Argument: 3 := 'Replace'
            }
            if ((type == 2))
            {
                Find = "N;";
                ReplaceWith = "";
                objDoc.Content.Find.Execute(ref Find, ref MatchCase, ref MatchWholeWord, ref MatchWildCards, ref MatchSoundsLike, ref MatchAllWordForms, ref Forward, ref Wrap, ref Format, ref ReplaceWith, ref Replace, ref MatchKashida, ref MatchDiacritics, ref MatchAlefHamza, ref MatchControl);
                // TODO: Labeled Arguments not supported. Argument: 1 := 'FindText'
                // TODO: Labeled Arguments not supported. Argument: 2 := 'ReplaceWith'
                // TODO: Labeled Arguments not supported. Argument: 3 := 'Replace'

                Find = "D;";
                ReplaceWith = "/";
                objDoc.Content.Find.Execute(ref Find, ref MatchCase, ref MatchWholeWord, ref MatchWildCards, ref MatchSoundsLike, ref MatchAllWordForms, ref Forward, ref Wrap, ref Format, ref ReplaceWith, ref Replace, ref MatchKashida, ref MatchDiacritics, ref MatchAlefHamza, ref MatchControl);
                // TODO: Labeled Arguments not supported. Argument: 1 := 'FindText'
                // TODO: Labeled Arguments not supported. Argument: 2 := 'ReplaceWith'
                // TODO: Labeled Arguments not supported. Argument: 3 := 'Replace'

                Find = "T;";
                ReplaceWith = "";
                objDoc.Content.Find.Execute(ref Find, ref MatchCase, ref MatchWholeWord, ref MatchWildCards, ref MatchSoundsLike, ref MatchAllWordForms, ref Forward, ref Wrap, ref Format, ref ReplaceWith, ref Replace, ref MatchKashida, ref MatchDiacritics, ref MatchAlefHamza, ref MatchControl);
                // TODO: Labeled Arguments not supported. Argument: 1 := 'FindText'
                // TODO: Labeled Arguments not supported. Argument: 2 := 'ReplaceWith'
                // TODO: Labeled Arguments not supported. Argument: 3 := 'Replace'
            }

            if ((type == 3))
            {
                // Find And Replace() some text.

                Find = "<TRANSFERFROM>";
                ReplaceWith = Asset;
                objDoc.Content.Find.Execute(ref Find, ref MatchCase, ref MatchWholeWord, ref MatchWildCards, ref MatchSoundsLike, ref MatchAllWordForms, ref Forward, ref Wrap, ref Format, ref ReplaceWith, ref Replace, ref MatchKashida, ref MatchDiacritics, ref MatchAlefHamza, ref MatchControl);
                // TODO: Labeled Arguments not supported. Argument: 1 := 'FindText'
                // TODO: Labeled Arguments not supported. Argument: 2 := 'ReplaceWith'
                // TODO: Labeled Arguments not supported. Argument: 3 := 'Replace'

                Find = "N;";
                ReplaceWith = "";
                objDoc.Content.Find.Execute(ref Find, ref MatchCase, ref MatchWholeWord, ref MatchWildCards, ref MatchSoundsLike, ref MatchAllWordForms, ref Forward, ref Wrap, ref Format, ref ReplaceWith, ref Replace, ref MatchKashida, ref MatchDiacritics, ref MatchAlefHamza, ref MatchControl);
                // TODO: Labeled Arguments not supported. Argument: 1 := 'FindText'
                // TODO: Labeled Arguments not supported. Argument: 2 := 'ReplaceWith'
                // TODO: Labeled Arguments not supported. Argument: 3 := 'Replace'

                Find = "D;";
                ReplaceWith = "";
                objDoc.Content.Find.Execute(ref Find, ref MatchCase, ref MatchWholeWord, ref MatchWildCards, ref MatchSoundsLike, ref MatchAllWordForms, ref Forward, ref Wrap, ref Format, ref ReplaceWith, ref Replace, ref MatchKashida, ref MatchDiacritics, ref MatchAlefHamza, ref MatchControl);
                // TODO: Labeled Arguments not supported. Argument: 1 := 'FindText'
                // TODO: Labeled Arguments not supported. Argument: 2 := 'ReplaceWith'
                // TODO: Labeled Arguments not supported. Argument: 3 := 'Replace'

                Find = "T;";
                ReplaceWith = "/";
                objDoc.Content.Find.Execute(ref Find, ref MatchCase, ref MatchWholeWord, ref MatchWildCards, ref MatchSoundsLike, ref MatchAllWordForms, ref Forward, ref Wrap, ref Format, ref ReplaceWith, ref Replace, ref MatchKashida, ref MatchDiacritics, ref MatchAlefHamza, ref MatchControl);
                // TODO: Labeled Arguments not supported. Argument: 1 := 'FindText'
                // TODO: Labeled Arguments not supported. Argument: 2 := 'ReplaceWith'
                // TODO: Labeled Arguments not supported. Argument: 3 := 'Replace'
            }
            else {


                // Find And Replace() some text.

                Find = "<ASSET>";
                ReplaceWith = Asset;
            objDoc.Content.Find.Execute(ref Find, ref MatchCase, ref MatchWholeWord, ref MatchWildCards, ref MatchSoundsLike, ref MatchAllWordForms, ref Forward, ref Wrap, ref Format, ref ReplaceWith, ref Replace, ref MatchKashida, ref MatchDiacritics, ref MatchAlefHamza, ref MatchControl);
                // TODO: Labeled Arguments not supported. Argument: 1 := 'FindText'
                // TODO: Labeled Arguments not supported. Argument: 2 := 'ReplaceWith'
                // TODO: Labeled Arguments not supported. Argument: 3 := 'Replace'
            }
            Find = "<SERIAL>";
            ReplaceWith = SN;
            // Find And Replace() some text.
            objDoc.Content.Find.Execute(ref Find, ref MatchCase, ref MatchWholeWord, ref MatchWildCards, ref MatchSoundsLike, ref MatchAllWordForms, ref Forward, ref Wrap, ref Format, ref ReplaceWith, ref Replace, ref MatchKashida, ref MatchDiacritics, ref MatchAlefHamza, ref MatchControl);
            // TODO: Labeled Arguments not supported. Argument: 1 := 'FindText'
            // TODO: Labeled Arguments not supported. Argument: 2 := 'ReplaceWith'
            // TODO: Labeled Arguments not supported. Argument: 3 := 'Replace'
            //  Find And Replace() some text.

            Find = "<ITEM>";
            ReplaceWith = description;
            objDoc.Content.Find.Execute(ref Find, ref MatchCase, ref MatchWholeWord, ref MatchWildCards, ref MatchSoundsLike, ref MatchAllWordForms, ref Forward, ref Wrap, ref Format, ref ReplaceWith, ref Replace, ref MatchKashida, ref MatchDiacritics, ref MatchAlefHamza, ref MatchControl);
            // TODO: Labeled Arguments not supported. Argument: 1 := 'FindText'
            // TODO: Labeled Arguments not supported. Argument: 2 := 'ReplaceWith'
            // TODO: Labeled Arguments not supported. Argument: 3 := 'Replace'
            if (((PO != "")
                        && (type != 3)))
            {

                Find = "<PO>";
                ReplaceWith = PO;
                //  Find And Replace() some text.
                objDoc.Content.Find.Execute(ref Find, ref MatchCase, ref MatchWholeWord, ref MatchWildCards, ref MatchSoundsLike, ref MatchAllWordForms, ref Forward, ref Wrap, ref Format, ref ReplaceWith, ref Replace, ref MatchKashida, ref MatchDiacritics, ref MatchAlefHamza, ref MatchControl);
                // TODO: Labeled Arguments not supported. Argument: 1 := 'FindText'
                // TODO: Labeled Arguments not supported. Argument: 2 := 'ReplaceWith'
                // TODO: Labeled Arguments not supported. Argument: 3 := 'Replace'
            }
            else if (((PO != "0")
                        && (type == 3)))
            {

                Find = "<TRANSFERTO";
                ReplaceWith = PO;
                //  Find And Replace() some text.
                objDoc.Content.Find.Execute(ref Find, ref MatchCase, ref MatchWholeWord, ref MatchWildCards, ref MatchSoundsLike, ref MatchAllWordForms, ref Forward, ref Wrap, ref Format, ref ReplaceWith, ref Replace, ref MatchKashida, ref MatchDiacritics, ref MatchAlefHamza, ref MatchControl);
                // TODO: Labeled Arguments not supported. Argument: 1 := 'FindText'
                // TODO: Labeled Arguments not supported. Argument: 2 := 'ReplaceWith'
                // TODO: Labeled Arguments not supported. Argument: 3 := 'Replace'
            }
            else {
                //  Find And Replace() some text.

                Find = "<PO>";
                ReplaceWith = "";
                objDoc.Content.Find.Execute(ref Find, ref MatchCase, ref MatchWholeWord, ref MatchWildCards, ref MatchSoundsLike, ref MatchAllWordForms, ref Forward, ref Wrap, ref Format, ref ReplaceWith, ref Replace, ref MatchKashida, ref MatchDiacritics, ref MatchAlefHamza, ref MatchControl);
                // TODO: Labeled Arguments not supported. Argument: 1 := 'FindText'
                // TODO: Labeled Arguments not supported. Argument: 2 := 'ReplaceWith'
                // TODO: Labeled Arguments not supported. Argument: 3 := 'Replace'
            }

            if ((transferto != "0"))
            {
                Find = "<TRANSFERTO";
                ReplaceWith = transferto;
                //  Find And Replace() some text.
                objDoc.Content.Find.Execute(ref Find, ref MatchCase, ref MatchWholeWord, ref MatchWildCards, ref MatchSoundsLike, ref MatchAllWordForms, ref Forward, ref Wrap, ref Format, ref ReplaceWith, ref Replace, ref MatchKashida, ref MatchDiacritics, ref MatchAlefHamza, ref MatchControl);
                // TODO: Labeled Arguments not supported. Argument: 1 := 'FindText'
                // TODO: Labeled Arguments not supported. Argument: 2 := 'ReplaceWith'
                // TODO: Labeled Arguments not supported. Argument: 3 := 'Replace'

                Find = "<TRANSFERFROM";
                ReplaceWith = transferfrom;
                //  Find And Replace() some text.
                objDoc.Content.Find.Execute(ref Find, ref MatchCase, ref MatchWholeWord, ref MatchWildCards, ref MatchSoundsLike, ref MatchAllWordForms, ref Forward, ref Wrap, ref Format, ref ReplaceWith, ref Replace, ref MatchKashida, ref MatchDiacritics, ref MatchAlefHamza, ref MatchControl);
                // TODO: Labeled Arguments not supported. Argument: 1 := 'FindText'
                // TODO: Labeled Arguments not supported. Argument: 2 := 'ReplaceWith'
                // TODO: Labeled Arguments not supported. Argument: 3 := 'Replace'
            }
            else {

                Find = "<TRANSFERTO>";
                ReplaceWith = "";
                //  Find And Replace() some text.
                objDoc.Content.Find.Execute(ref Find, ref MatchCase, ref MatchWholeWord, ref MatchWildCards, ref MatchSoundsLike, ref MatchAllWordForms, ref Forward, ref Wrap, ref Format, ref ReplaceWith, ref Replace, ref MatchKashida, ref MatchDiacritics, ref MatchAlefHamza, ref MatchControl);
                // TODO: Labeled Arguments not supported. Argument: 1 := 'FindText'
                // TODO: Labeled Arguments not supported. Argument: 2 := 'ReplaceWith'
                // TODO: Labeled Arguments not supported. Argument: 3 := 'Replace'
                //  Find And Replace() some text.

                Find = "<TRANSFERFROM>";
                ReplaceWith = "";
                objDoc.Content.Find.Execute(ref Find, ref MatchCase, ref MatchWholeWord, ref MatchWildCards, ref MatchSoundsLike, ref MatchAllWordForms, ref Forward, ref Wrap, ref Format, ref ReplaceWith, ref Replace, ref MatchKashida, ref MatchDiacritics, ref MatchAlefHamza, ref MatchControl);
                // TODO: Labeled Arguments not supported. Argument: 1 := 'FindText'
                // TODO: Labeled Arguments not supported. Argument: 2 := 'ReplaceWith'
                // TODO: Labeled Arguments not supported. Argument: 3 := 'Replace'
            }
            toolStripProgressBar1.Value = 80;
            objWordApp.PrintOut();
            // Save and close the document.
            objDoc.SaveAs("E:\\C#\\final.docx");
            //   objWordApp.Documents.Close(Word.WdSaveOptions.wdDoNotSaveChanges)
            objWordApp.Quit();
            objWordApp = null;
        }

        private void dategrab()//Called to grab the date and set the label
        {
            
            DateLabel.Text = dateran;//Sets the date in the label
            
        }

        private void formatmessage()//Creates a TXT file with all of the information
        {
            string text;
            if ((type == 1))
            {
                text = ("Received "
                            + (description
                            + (Environment.NewLine
                            + (Environment.NewLine + ("Description "
                            + (description
                            + (Environment.NewLine + ("Asset "
                            + (Asset
                            + (Environment.NewLine + ("Serial Numer " + SN)))))))))));
                if ((WO != "0"))
                {
                    text = (text+ (Environment.NewLine + ("Work Order " + WO)));
                }

                if ((MAC != "0"))
                {
                    text = (text+ (Environment.NewLine + ("MAC " + MAC)));
                }

                if ((PO != "0"))
                {
                    text = (text+ (Environment.NewLine + ("Purchase Order " + PO)));
                }

                text = text + Environment.NewLine + Environment.NewLine;
                System.IO.File.AppendAllText("D:\\C#\\WriteText.txt", text);
                text = null;
                toolStripProgressBar1.Value = 50;
            }
            else if ((type == 2))
            {
                text = ("Disposed "
                            + (description
                            + (Environment.NewLine
                            + (Environment.NewLine + ("Description "
                            + (description
                            + (Environment.NewLine + ("Asset "
                            + (Asset
                            + (Environment.NewLine + ("Serial Numer " + SN)))))))))));
                if ((WO != "0"))
                {
                    text = (text
                                + (Environment.NewLine + ("Work Order " + WO)));
                }

                if ((MAC != "0"))
                {
                    text = (text
                                + (Environment.NewLine + ("MAC " + MAC)));
                }

                if ((PO != "0"))
                {
                    text = (text
                                + (Environment.NewLine + ("Purchase Order " + PO)));
                }
                text = text + Environment.NewLine + Environment.NewLine;
                System.IO.File.AppendAllText("V:\\C#\\WriteText.txt", text);
                System.Diagnostics.Process.Start("V:\\C#\\WriteText.txt");
                text = null;
                toolStripProgressBar1.Value = 50;
            }
            else if ((type == 3))
            {
                text = ("Transfered "
                            + (description
                            + (Environment.NewLine
                            + (Environment.NewLine + ("Description "
                            + (description
                            + (Environment.NewLine + ("Asset "
                            + (Asset
                            + (Environment.NewLine + ("Serial Numer " + SN)))))))))));
                if ((WO != "0"))
                {
                    text = (text
                                + (Environment.NewLine + ("Work Order " + WO)));
                }

                if ((MAC != "0"))
                {
                    text = (text
                                + (Environment.NewLine + ("MAC " + MAC)));
                }

                if ((PO != "0"))
                {
                    text = (text+ (Environment.NewLine + ("Purchase Order " + PO)));
                }
                text = text + Environment.NewLine + Environment.NewLine;
                System.IO.File.AppendAllText("V:\\C#\\WriteText.txt", text);
                text = null;
            }
            toolStripProgressBar1.Value = 50;


        }

        private void ComboBox1_KeyDown(object sender, KeyEventArgs e)//Called when the combo box hears a key
        {
            if (((e.KeyCode == Keys.Enter)|| (e.KeyCode == Keys.Return)))
                //^^Listens for Enter or Return key
            {
                description = ComboBox1.Text;
                TextBox1.Focus();
                Label6.Visible = true;
                Label6.Text = description;
                ComboBox1.Text = "Description";
            }

        }

        private void ComboBox1_Click(object sender, EventArgs e)//Called when focus is changed to the Combo box
        {
            ComboBox1.Text = "";//On clcik clear the text
        }

        private void ComboBox1_Leave(object sender, EventArgs e)//Called if the focus leaves the combobox
        {
            ComboBox1.Text = "Description";//Sets the text back to description
        }

        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)//Called if the combox item is changed
        {
            description = ComboBox1.SelectedItem.ToString();//Converts selected to string and adds to varible
            Label6.Text = description;//sets the label 
            Label6.Visible = true;//Makes label visible
            ComboBox1.Text = "Description";//Sets text back to description
        }

        private void ComboBox1_Enter(object sender, EventArgs e)//Called if combobox hears the enter key
        {
            ComboBox1.Text = "";//Set text to null
        }

        private void timer1_tick(object sender, EventArgs e)//Clock
        {
            timelabel.Text = DateTime.Now.ToString("HH:mm");//Sets time in label
            if (((DateTime.Now.ToString("HH:mm:ss") == "00:00:00")//Checks if midnight struck
                        && (passing == false)))
            {
                passing = true;//Informs the date was grabbed
               

            }
        }
        private void pulldata()
        {
            connection = new MySqlConnection(ConnectionString);
             connection.Open();
            string cmd = "Select * FROM assettracker.AssetTable";
            adapter = new MySqlDataAdapter(cmd, connection);
            System.Data.DataTable datatable = new System.Data.DataTable();
            {
               
                DataSet Allasset = new DataSet();
                Allasset.ToString();
                adapter.Fill(datatable);
                listBox1.DataSource = Allasset;
                
                    }

            
        }
        private void pushdata()
        {
            string addquery = "Insert into assettracker.AssetTable Values (@Asset , @Type, @Description, @Purchase_Order, @Serial, @Work_Order, @MAC, @Department, @User_First, @User_Last, @Employee_ID, @Date, @Modified, @Deleted, @Asset_To, @Asset_From)";
            connection = new MySqlConnection(ConnectionString);
            command = connection.CreateCommand();
            {
                if (transferfrom == "0")
                {
                    transferfrom = "NULL";
                }
                if (transferto == "0")
                {
                    transferfrom = "NULL";
                }
                if (department == "0")
                {
                    transferfrom = "NULL";
                }

                try
                {
                    command.CommandText = addquery;
                    command.Parameters.AddWithValue("@User_First", user);
                    command.Parameters.AddWithValue("@User_Last", user);
                    command.Parameters.AddWithValue("@Employee_ID", user);

                    command.Parameters.AddWithValue("@Type", assettype);
                    command.Parameters.AddWithValue("@Asset", Asset);
                    command.Parameters.AddWithValue("@Description", description);
                    command.Parameters.AddWithValue("@Purchase_Order", user);
                    command.Parameters.AddWithValue("@Department", department);
                    command.Parameters.AddWithValue("@Asset_To", transferto);
                    command.Parameters.AddWithValue("@Asset_From", transferfrom);
                    command.Parameters.AddWithValue("@Date", datecomplete);
                    command.Parameters.AddWithValue("@Serial", SN);
                    command.Parameters.AddWithValue("@Work_Order", WO);
                    command.Parameters.AddWithValue("@MAC", MAC);

                    command.Parameters.AddWithValue("@Modified", "TEMP");
                    command.Parameters.AddWithValue("@Deleted", "TEMP");


                    command.ExecuteNonQuery();
                }
                catch (Exception)
                {
                    throw;
                }

                if (connection.State == ConnectionState.Open)
                {
                    connection.Close();
                }
            }


        }



        private void Label6_MouseHover(object sender, EventArgs e)
        {
            int xtemp = (Control.MousePosition.X);
            int ytemp = (Control.MousePosition.Y);
            System.Drawing.Point y = PointToScreen(new System.Drawing.Point(ytemp));
            System.Drawing.Point x = PointToScreen(new System.Drawing.Point(xtemp));
            int xconvert = Convert.ToInt16(x.X);
            int yconvert = Convert.ToInt16(y.Y);
            ToolTip descrptool = new ToolTip();
            descrptool.InitialDelay = 0;
            descrptool.ShowAlways = true;
            descrptool.UseFading = true;
            descrptool.UseAnimation = true;
            descrptool.IsBalloon = false;
            descrptool.ToolTipTitle = "Description";
            descrptool.Show("Double Click for edit mode",TextBox1,xconvert,yconvert, 2000);

            


        }

        private void toolStripDropDownButton2_Click(object sender, EventArgs e)
        {

        }

        private void pullToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void Porder_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            scannedtemp = "";
            PO = "0";
            Porder.Text = "";
            Porder.BackColor = Color.FromKnownColor(KnownColor.Menu);
        }

        private void assettag_MouseDoubleClick(object sender, MouseEventArgs e)
        {

        }

        private void Label6_MouseDoubleClick(object sender, MouseEventArgs e)
        {

        }

        private void Serial_MouseDoubleClick(object sender, MouseEventArgs e)
        {

        }

        private void Worder_MouseDoubleClick(object sender, MouseEventArgs e)
        {

        }

        private void labeldepartment_MouseDoubleClick(object sender, MouseEventArgs e)
        {

        }

        private void Maca_MouseDoubleClick(object sender, MouseEventArgs e)
        {

        }

        

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start("notepad.exe", "\\C#/Asset/WindowsFormsApplication1/Device Description");
        }

        private void editToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // var fileStream = new FileStream("\\C#/Asset/WindowsFormsApplication1/Device Description", FileMode.Open);
        }

        private void settingsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Change where to look for the file
        }
    }
}
