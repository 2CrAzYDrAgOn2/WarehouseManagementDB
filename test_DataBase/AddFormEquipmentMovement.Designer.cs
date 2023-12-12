namespace test_DataBase
{
    partial class AddFormEquipmentMovement
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
            this.label1 = new System.Windows.Forms.Label();
            this.labelTitle = new System.Windows.Forms.Label();
            this.buttonSave = new System.Windows.Forms.Button();
            this.label22 = new System.Windows.Forms.Label();
            this.textBoxQuantinityEquipmentMovement = new System.Windows.Forms.TextBox();
            this.label18 = new System.Windows.Forms.Label();
            this.label20 = new System.Windows.Forms.Label();
            this.textBoxMovementType = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.textBoxEquipmentIDEquipmentMovement = new System.Windows.Forms.TextBox();
            this.textBoxMovementDate = new System.Windows.Forms.DateTimePicker();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(207, 35);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(250, 21);
            this.label1.TabIndex = 38;
            this.label1.Text = "Передвижение оборудования";
            // 
            // labelTitle
            // 
            this.labelTitle.AutoSize = true;
            this.labelTitle.Font = new System.Drawing.Font("Segoe UI", 14.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.labelTitle.Location = new System.Drawing.Point(206, 10);
            this.labelTitle.Name = "labelTitle";
            this.labelTitle.Size = new System.Drawing.Size(175, 25);
            this.labelTitle.TabIndex = 37;
            this.labelTitle.Text = "Создание записи:";
            // 
            // buttonSave
            // 
            this.buttonSave.Location = new System.Drawing.Point(281, 662);
            this.buttonSave.Name = "buttonSave";
            this.buttonSave.Size = new System.Drawing.Size(202, 56);
            this.buttonSave.TabIndex = 28;
            this.buttonSave.Text = "Сохранить";
            this.buttonSave.UseVisualStyleBackColor = true;
            this.buttonSave.Click += new System.EventHandler(this.ButtonSave_Click);
            // 
            // label22
            // 
            this.label22.AutoSize = true;
            this.label22.Location = new System.Drawing.Point(170, 485);
            this.label22.Name = "label22";
            this.label22.Size = new System.Drawing.Size(69, 13);
            this.label22.TabIndex = 51;
            this.label22.Text = "Количество:";
            // 
            // textBoxQuantinityEquipmentMovement
            // 
            this.textBoxQuantinityEquipmentMovement.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBoxQuantinityEquipmentMovement.Location = new System.Drawing.Point(245, 473);
            this.textBoxQuantinityEquipmentMovement.Name = "textBoxQuantinityEquipmentMovement";
            this.textBoxQuantinityEquipmentMovement.Size = new System.Drawing.Size(391, 33);
            this.textBoxQuantinityEquipmentMovement.TabIndex = 50;
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(133, 446);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(106, 13);
            this.label18.TabIndex = 49;
            this.label18.Text = "Тип передвижения:";
            // 
            // label20
            // 
            this.label20.AutoSize = true;
            this.label20.Location = new System.Drawing.Point(126, 407);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(113, 13);
            this.label20.TabIndex = 48;
            this.label20.Text = "Дата передвижения:";
            // 
            // textBoxMovementType
            // 
            this.textBoxMovementType.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBoxMovementType.Location = new System.Drawing.Point(245, 434);
            this.textBoxMovementType.Name = "textBoxMovementType";
            this.textBoxMovementType.Size = new System.Drawing.Size(391, 33);
            this.textBoxMovementType.TabIndex = 47;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(121, 368);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(118, 13);
            this.label3.TabIndex = 45;
            this.label3.Text = "Номер оборудования:";
            // 
            // textBoxEquipmentIDEquipmentMovement
            // 
            this.textBoxEquipmentIDEquipmentMovement.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBoxEquipmentIDEquipmentMovement.Location = new System.Drawing.Point(245, 356);
            this.textBoxEquipmentIDEquipmentMovement.Name = "textBoxEquipmentIDEquipmentMovement";
            this.textBoxEquipmentIDEquipmentMovement.Size = new System.Drawing.Size(391, 33);
            this.textBoxEquipmentIDEquipmentMovement.TabIndex = 44;
            // 
            // textBoxMovementDate
            // 
            this.textBoxMovementDate.Font = new System.Drawing.Font("Segoe UI", 14.25F);
            this.textBoxMovementDate.Location = new System.Drawing.Point(245, 395);
            this.textBoxMovementDate.Name = "textBoxMovementDate";
            this.textBoxMovementDate.Size = new System.Drawing.Size(391, 33);
            this.textBoxMovementDate.TabIndex = 52;
            // 
            // AddFormEquipmentMovement
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(768, 729);
            this.Controls.Add(this.textBoxMovementDate);
            this.Controls.Add(this.label22);
            this.Controls.Add(this.textBoxQuantinityEquipmentMovement);
            this.Controls.Add(this.label18);
            this.Controls.Add(this.label20);
            this.Controls.Add(this.textBoxMovementType);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.textBoxEquipmentIDEquipmentMovement);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.labelTitle);
            this.Controls.Add(this.buttonSave);
            this.Name = "AddFormEquipmentMovement";
            this.Text = "Добавить передвижение оборудования";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label labelTitle;
        private System.Windows.Forms.Button buttonSave;
        private System.Windows.Forms.Label label22;
        private System.Windows.Forms.TextBox textBoxQuantinityEquipmentMovement;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.Label label20;
        private System.Windows.Forms.TextBox textBoxMovementType;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBoxEquipmentIDEquipmentMovement;
        private System.Windows.Forms.DateTimePicker textBoxMovementDate;
    }
}