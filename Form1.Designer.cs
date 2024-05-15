namespace CorreosAutomaticos
{
    partial class Form1
    {
        /// <summary>
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de Windows Forms

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.BtnAdjuntar = new System.Windows.Forms.Button();
            this.BtnEnviar = new System.Windows.Forms.Button();
            this.TxtLblAdj = new System.Windows.Forms.Label();
            this.BtnActualizar = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // BtnAdjuntar
            // 
            this.BtnAdjuntar.Location = new System.Drawing.Point(114, 252);
            this.BtnAdjuntar.Name = "BtnAdjuntar";
            this.BtnAdjuntar.Size = new System.Drawing.Size(75, 41);
            this.BtnAdjuntar.TabIndex = 0;
            this.BtnAdjuntar.Text = "Adjuntar";
            this.BtnAdjuntar.UseVisualStyleBackColor = true;
            this.BtnAdjuntar.Click += new System.EventHandler(this.BtnAdjuntar_Click);
            // 
            // BtnEnviar
            // 
            this.BtnEnviar.Location = new System.Drawing.Point(348, 252);
            this.BtnEnviar.Name = "BtnEnviar";
            this.BtnEnviar.Size = new System.Drawing.Size(77, 41);
            this.BtnEnviar.TabIndex = 1;
            this.BtnEnviar.Text = "Enviar";
            this.BtnEnviar.UseVisualStyleBackColor = true;
            this.BtnEnviar.Click += new System.EventHandler(this.BtnEnviar_Click);
            // 
            // TxtLblAdj
            // 
            this.TxtLblAdj.AutoSize = true;
            this.TxtLblAdj.Location = new System.Drawing.Point(137, 190);
            this.TxtLblAdj.Name = "TxtLblAdj";
            this.TxtLblAdj.Size = new System.Drawing.Size(16, 13);
            this.TxtLblAdj.TabIndex = 2;
            this.TxtLblAdj.Text = "...";
            // 
            // BtnActualizar
            // 
            this.BtnActualizar.Location = new System.Drawing.Point(236, 261);
            this.BtnActualizar.Name = "BtnActualizar";
            this.BtnActualizar.Size = new System.Drawing.Size(75, 23);
            this.BtnActualizar.TabIndex = 3;
            this.BtnActualizar.Text = "Actualizar";
            this.BtnActualizar.UseVisualStyleBackColor = true;
            this.BtnActualizar.Click += new System.EventHandler(this.BtnActualizar_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToOrderColumns = true;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(27, 12);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(514, 166);
            this.dataGridView1.TabIndex = 4;
            // 
            // timer1
            // 
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(572, 305);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.BtnActualizar);
            this.Controls.Add(this.TxtLblAdj);
            this.Controls.Add(this.BtnEnviar);
            this.Controls.Add(this.BtnAdjuntar);
            this.Name = "Form1";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button BtnAdjuntar;
        private System.Windows.Forms.Button BtnEnviar;
        private System.Windows.Forms.Label TxtLblAdj;
        private System.Windows.Forms.Button BtnActualizar;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Timer timer1;
    }
}

