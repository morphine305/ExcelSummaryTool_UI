namespace ExcelPlotTool
{
    partial class Form1
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置受控資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.uiTableLayoutPanel1 = new Sunny.UI.UITableLayoutPanel();
            this.uiLabel1 = new Sunny.UI.UILabel();
            this.Folder_path_tb = new Sunny.UI.UITextBox();
            this.start_bt = new Sunny.UI.UIButton();
            this.uiTableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // uiTableLayoutPanel1
            // 
            this.uiTableLayoutPanel1.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.InsetDouble;
            this.uiTableLayoutPanel1.ColumnCount = 2;
            this.uiTableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 19.10828F));
            this.uiTableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 80.89172F));
            this.uiTableLayoutPanel1.Controls.Add(this.uiLabel1, 0, 0);
            this.uiTableLayoutPanel1.Controls.Add(this.Folder_path_tb, 1, 0);
            this.uiTableLayoutPanel1.Location = new System.Drawing.Point(32, 72);
            this.uiTableLayoutPanel1.Name = "uiTableLayoutPanel1";
            this.uiTableLayoutPanel1.RowCount = 1;
            this.uiTableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.uiTableLayoutPanel1.Size = new System.Drawing.Size(471, 43);
            this.uiTableLayoutPanel1.TabIndex = 0;
            this.uiTableLayoutPanel1.TagString = null;
            // 
            // uiLabel1
            // 
            this.uiLabel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.uiLabel1.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.uiLabel1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(48)))), ((int)(((byte)(48)))), ((int)(((byte)(48)))));
            this.uiLabel1.Location = new System.Drawing.Point(6, 3);
            this.uiLabel1.Name = "uiLabel1";
            this.uiLabel1.Size = new System.Drawing.Size(82, 37);
            this.uiLabel1.TabIndex = 0;
            this.uiLabel1.Text = "uiLabel1";
            this.uiLabel1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // Folder_path_tb
            // 
            this.Folder_path_tb.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.Folder_path_tb.Dock = System.Windows.Forms.DockStyle.Fill;
            this.Folder_path_tb.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.Folder_path_tb.Location = new System.Drawing.Point(98, 8);
            this.Folder_path_tb.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Folder_path_tb.MinimumSize = new System.Drawing.Size(1, 16);
            this.Folder_path_tb.Name = "Folder_path_tb";
            this.Folder_path_tb.Padding = new System.Windows.Forms.Padding(5);
            this.Folder_path_tb.ShowText = false;
            this.Folder_path_tb.Size = new System.Drawing.Size(366, 27);
            this.Folder_path_tb.TabIndex = 2;
            this.Folder_path_tb.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft;
            this.Folder_path_tb.Watermark = "";
            this.Folder_path_tb.Click += new System.EventHandler(this.SelectFolder_Click);
            // 
            // start_bt
            // 
            this.start_bt.Cursor = System.Windows.Forms.Cursors.Hand;
            this.start_bt.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.start_bt.Location = new System.Drawing.Point(543, 75);
            this.start_bt.MinimumSize = new System.Drawing.Size(1, 1);
            this.start_bt.Name = "start_bt";
            this.start_bt.Size = new System.Drawing.Size(100, 35);
            this.start_bt.TabIndex = 1;
            this.start_bt.Text = "uiButton1";
            this.start_bt.TipsFont = new System.Drawing.Font("新細明體", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.start_bt.Click += new System.EventHandler(this.start_bt_Click);
            // 
            // Form1
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.start_bt);
            this.Controls.Add(this.uiTableLayoutPanel1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ZoomScaleRect = new System.Drawing.Rectangle(15, 15, 800, 450);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.uiTableLayoutPanel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private Sunny.UI.UITableLayoutPanel uiTableLayoutPanel1;
        private Sunny.UI.UILabel uiLabel1;
        private Sunny.UI.UITextBox Folder_path_tb;
        private Sunny.UI.UIButton start_bt;
    }
}

