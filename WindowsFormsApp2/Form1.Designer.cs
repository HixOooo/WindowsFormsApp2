using System.Windows.Forms;

namespace WindowsFormsApp1
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.lblTitle = new System.Windows.Forms.Label();
            this.txtTitle = new System.Windows.Forms.TextBox();
            this.lblAuthor = new System.Windows.Forms.Label();
            this.txtAuthor = new System.Windows.Forms.TextBox();
            this.lblYear = new System.Windows.Forms.Label();
            this.txtYear = new System.Windows.Forms.TextBox();
            this.lstBooks = new System.Windows.Forms.ListBox();
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnDelete = new System.Windows.Forms.Button();
            this.btnSearchTitle = new System.Windows.Forms.Button();
            this.btnSearchAuthor = new System.Windows.Forms.Button();
            this.btnExport = new System.Windows.Forms.Button();
            this.btnImport = new System.Windows.Forms.Button();
            this.btnShowAllBooks = new System.Windows.Forms.Button();
            this.btnExportToDocx = new System.Windows.Forms.Button();
            this.btnExportToPdf = new System.Windows.Forms.Button();
            this.btnImportFromPdf = new System.Windows.Forms.Button();
            this.btnImportFromDocx = new System.Windows.Forms.Button();
            this.btnConvert = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lblTitle
            // 
            this.lblTitle.AutoSize = true;
            this.lblTitle.Location = new System.Drawing.Point(12, 15);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(60, 13);
            this.lblTitle.TabIndex = 0;
            this.lblTitle.Text = "Название:";
            // 
            // txtTitle
            // 
            this.txtTitle.Location = new System.Drawing.Point(84, 12);
            this.txtTitle.Name = "txtTitle";
            this.txtTitle.Size = new System.Drawing.Size(109, 20);
            this.txtTitle.TabIndex = 1;
            // 
            // lblAuthor
            // 
            this.lblAuthor.AutoSize = true;
            this.lblAuthor.Location = new System.Drawing.Point(12, 45);
            this.lblAuthor.Name = "lblAuthor";
            this.lblAuthor.Size = new System.Drawing.Size(40, 13);
            this.lblAuthor.TabIndex = 2;
            this.lblAuthor.Text = "Автор:";
            // 
            // txtAuthor
            // 
            this.txtAuthor.Location = new System.Drawing.Point(84, 42);
            this.txtAuthor.Name = "txtAuthor";
            this.txtAuthor.Size = new System.Drawing.Size(109, 20);
            this.txtAuthor.TabIndex = 3;
            // 
            // lblYear
            // 
            this.lblYear.AutoSize = true;
            this.lblYear.Location = new System.Drawing.Point(12, 75);
            this.lblYear.Name = "lblYear";
            this.lblYear.Size = new System.Drawing.Size(28, 13);
            this.lblYear.TabIndex = 4;
            this.lblYear.Text = "Год:";
            // 
            // txtYear
            // 
            this.txtYear.Location = new System.Drawing.Point(84, 72);
            this.txtYear.Name = "txtYear";
            this.txtYear.Size = new System.Drawing.Size(109, 20);
            this.txtYear.TabIndex = 5;
            // 
            // lstBooks
            // 
            this.lstBooks.FormattingEnabled = true;
            this.lstBooks.Location = new System.Drawing.Point(15, 104);
            this.lstBooks.Name = "lstBooks";
            this.lstBooks.Size = new System.Drawing.Size(370, 290);
            this.lstBooks.TabIndex = 6;
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(199, 10);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 23);
            this.btnAdd.TabIndex = 7;
            this.btnAdd.Text = "Добавить";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnDelete
            // 
            this.btnDelete.Location = new System.Drawing.Point(199, 40);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(75, 23);
            this.btnDelete.TabIndex = 8;
            this.btnDelete.Text = "Удалить";
            this.btnDelete.UseVisualStyleBackColor = true;
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // btnSearchTitle
            // 
            this.btnSearchTitle.Location = new System.Drawing.Point(280, 24);
            this.btnSearchTitle.Name = "btnSearchTitle";
            this.btnSearchTitle.Size = new System.Drawing.Size(116, 23);
            this.btnSearchTitle.TabIndex = 9;
            this.btnSearchTitle.Text = "Поиск по названию";
            this.btnSearchTitle.UseVisualStyleBackColor = true;
            this.btnSearchTitle.Click += new System.EventHandler(this.btnSearchTitle_Click);
            // 
            // btnSearchAuthor
            // 
            this.btnSearchAuthor.Location = new System.Drawing.Point(280, 55);
            this.btnSearchAuthor.Name = "btnSearchAuthor";
            this.btnSearchAuthor.Size = new System.Drawing.Size(116, 23);
            this.btnSearchAuthor.TabIndex = 10;
            this.btnSearchAuthor.Text = "Поиск по автору";
            this.btnSearchAuthor.UseVisualStyleBackColor = true;
            this.btnSearchAuthor.Click += new System.EventHandler(this.btnSearchAuthor_Click);
            // 
            // btnExport
            // 
            this.btnExport.Location = new System.Drawing.Point(199, 70);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(75, 23);
            this.btnExport.TabIndex = 11;
            this.btnExport.Text = "Экспорт";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // btnImport
            // 
            this.btnImport.Location = new System.Drawing.Point(391, 104);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(115, 25);
            this.btnImport.TabIndex = 12;
            this.btnImport.Text = "Импорт";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // btnShowAllBooks
            // 
            this.btnShowAllBooks.Location = new System.Drawing.Point(402, 24);
            this.btnShowAllBooks.Name = "btnShowAllBooks";
            this.btnShowAllBooks.Size = new System.Drawing.Size(105, 54);
            this.btnShowAllBooks.TabIndex = 13;
            this.btnShowAllBooks.Text = "Показать все книги";
            this.btnShowAllBooks.UseVisualStyleBackColor = true;
            this.btnShowAllBooks.Click += new System.EventHandler(this.btnShowAllBooks_Click);
            // 
            // btnExportToDocx
            // 
            this.btnExportToDocx.Location = new System.Drawing.Point(391, 233);
            this.btnExportToDocx.Name = "btnExportToDocx";
            this.btnExportToDocx.Size = new System.Drawing.Size(116, 30);
            this.btnExportToDocx.TabIndex = 0;
            this.btnExportToDocx.Text = "Экспорт в DOCX";
            this.btnExportToDocx.UseVisualStyleBackColor = true;
            this.btnExportToDocx.Click += new System.EventHandler(this.btnExportToDocx_Click);
            // 
            // btnExportToPdf
            // 
            this.btnExportToPdf.Location = new System.Drawing.Point(391, 197);
            this.btnExportToPdf.Name = "btnExportToPdf";
            this.btnExportToPdf.Size = new System.Drawing.Size(116, 30);
            this.btnExportToPdf.TabIndex = 0;
            this.btnExportToPdf.Text = "Экспорт в PDF";
            this.btnExportToPdf.UseVisualStyleBackColor = true;
            this.btnExportToPdf.Click += new System.EventHandler(this.btnExportToPdf_Click);
            // 
            // btnImportFromPdf
            // 
            this.btnImportFromPdf.Location = new System.Drawing.Point(391, 135);
            this.btnImportFromPdf.Name = "btnImportFromPdf";
            this.btnImportFromPdf.Size = new System.Drawing.Size(115, 24);
            this.btnImportFromPdf.TabIndex = 0;
            this.btnImportFromPdf.Text = "Импорт из PDF";
            this.btnImportFromPdf.UseVisualStyleBackColor = true;
            this.btnImportFromPdf.Click += new System.EventHandler(this.btnImportFromPdf_Click);
            // 
            // btnImportFromDocx
            // 
            this.btnImportFromDocx.Location = new System.Drawing.Point(391, 165);
            this.btnImportFromDocx.Name = "btnImportFromDocx";
            this.btnImportFromDocx.Size = new System.Drawing.Size(115, 24);
            this.btnImportFromDocx.TabIndex = 1;
            this.btnImportFromDocx.Text = "Импорт из DOCX";
            this.btnImportFromDocx.UseVisualStyleBackColor = true;
            this.btnImportFromDocx.Click += new System.EventHandler(this.btnImportFromDocx_Click);
            // 
            // btnConvert
            // 
            this.btnConvert.Location = new System.Drawing.Point(391, 328);
            this.btnConvert.Name = "btnConvert";
            this.btnConvert.Size = new System.Drawing.Size(105, 66);
            this.btnConvert.TabIndex = 0;
            this.btnConvert.Text = "Конвертировать файл";
            this.btnConvert.UseVisualStyleBackColor = true;
            this.btnConvert.Click += new System.EventHandler(this.btnConvert_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(511, 406);
            this.Controls.Add(this.btnShowAllBooks);
            this.Controls.Add(this.btnImport);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.btnSearchAuthor);
            this.Controls.Add(this.btnSearchTitle);
            this.Controls.Add(this.btnDelete);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.lstBooks);
            this.Controls.Add(this.txtYear);
            this.Controls.Add(this.lblYear);
            this.Controls.Add(this.txtAuthor);
            this.Controls.Add(this.lblAuthor);
            this.Controls.Add(this.txtTitle);
            this.Controls.Add(this.lblTitle);
            this.Controls.Add(this.btnExportToDocx);
            this.Controls.Add(this.btnExportToPdf);
            this.Controls.Add(this.btnImportFromDocx);
            this.Controls.Add(this.btnImportFromPdf);
            this.Controls.Add(this.btnConvert);
            this.Name = "Form1";
            this.Text = "Менеджер Книг";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.TextBox txtTitle;
        private System.Windows.Forms.Label lblAuthor;
        private System.Windows.Forms.TextBox txtAuthor;
        private System.Windows.Forms.Label lblYear;
        private System.Windows.Forms.TextBox txtYear;
        private System.Windows.Forms.ListBox lstBooks;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Button btnDelete;
        private System.Windows.Forms.Button btnSearchTitle;
        private System.Windows.Forms.Button btnSearchAuthor;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.Button btnImport;
        private System.Windows.Forms.Button btnShowAllBooks;
        private System.Windows.Forms.Button btnExportToDocx;
        private System.Windows.Forms.Button btnExportToPdf;
        private System.Windows.Forms.Button btnImportFromPdf;
        private System.Windows.Forms.Button btnImportFromDocx;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.Button btnConvert;
    }
}

