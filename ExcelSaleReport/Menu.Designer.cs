namespace ExcelSaleReport
{
    partial class Menu
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.b_hour = new System.Windows.Forms.Button();
            this.b_day = new System.Windows.Forms.Button();
            this.b_supplier = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // b_hour
            // 
            this.b_hour.Dock = System.Windows.Forms.DockStyle.Top;
            this.b_hour.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.b_hour.Location = new System.Drawing.Point(0, 0);
            this.b_hour.Name = "b_hour";
            this.b_hour.Size = new System.Drawing.Size(316, 76);
            this.b_hour.TabIndex = 0;
            this.b_hour.Text = "Realization by hour";
            this.b_hour.UseVisualStyleBackColor = true;
            // 
            // b_day
            // 
            this.b_day.Dock = System.Windows.Forms.DockStyle.Top;
            this.b_day.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.b_day.Location = new System.Drawing.Point(0, 76);
            this.b_day.Name = "b_day";
            this.b_day.Size = new System.Drawing.Size(316, 76);
            this.b_day.TabIndex = 1;
            this.b_day.Text = "Realization by day";
            this.b_day.UseVisualStyleBackColor = true;
            // 
            // b_supplier
            // 
            this.b_supplier.Dock = System.Windows.Forms.DockStyle.Top;
            this.b_supplier.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.b_supplier.Location = new System.Drawing.Point(0, 152);
            this.b_supplier.Name = "b_supplier";
            this.b_supplier.Size = new System.Drawing.Size(316, 76);
            this.b_supplier.TabIndex = 2;
            this.b_supplier.Text = "Realization by supplier";
            this.b_supplier.UseVisualStyleBackColor = true;
            // 
            // Menu
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(316, 228);
            this.Controls.Add(this.b_supplier);
            this.Controls.Add(this.b_day);
            this.Controls.Add(this.b_hour);
            this.Name = "Menu";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Excel sale report";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button b_hour;
        private System.Windows.Forms.Button b_day;
        private System.Windows.Forms.Button b_supplier;
    }
}

