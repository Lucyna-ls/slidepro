using System;using System.Drawing;namespace PowerPointAddIn2{    partial class UserControl1    {
        // Existing declarations
        private System.Windows.Forms.Button buttonGenerate;        private System.Windows.Forms.Button buttonLoadMore; // Declare the Load More button
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanelSlides;        private System.Windows.Forms.PictureBox pictureBoxLoader;



        #region Component Designer generated code
        private void InitializeComponent()        {            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(UserControl1));
            this.flowLayoutPanelSlides = new System.Windows.Forms.FlowLayoutPanel();
            this.pictureBoxLoader = new System.Windows.Forms.PictureBox();
            this.buttonGenerate = new System.Windows.Forms.Button();
            this.buttonLoadMore = new System.Windows.Forms.Button();
            this.flowLayoutPanelSlides.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxLoader)).BeginInit();
            this.SuspendLayout();
            // 
            // flowLayoutPanelSlides
            // 
            this.flowLayoutPanelSlides.AutoScroll = true;
            this.flowLayoutPanelSlides.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(248)))), ((int)(((byte)(244)))), ((int)(((byte)(244)))));
            this.flowLayoutPanelSlides.Controls.Add(this.pictureBoxLoader);
            this.flowLayoutPanelSlides.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.flowLayoutPanelSlides.Location = new System.Drawing.Point(25, 51);
            this.flowLayoutPanelSlides.Name = "flowLayoutPanelSlides";
            this.flowLayoutPanelSlides.Padding = new System.Windows.Forms.Padding(5);
            this.flowLayoutPanelSlides.Size = new System.Drawing.Size(380, 611);
            this.flowLayoutPanelSlides.TabIndex = 0;
            this.flowLayoutPanelSlides.WrapContents = false;
            this.flowLayoutPanelSlides.Paint += new System.Windows.Forms.PaintEventHandler(this.flowLayoutPanelSlides_Paint);
            // 
            // pictureBoxLoader
            // 
            this.pictureBoxLoader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(248)))), ((int)(((byte)(244)))), ((int)(((byte)(244)))));
            this.pictureBoxLoader.Image = ((System.Drawing.Image)(resources.GetObject("pictureBoxLoader.Image")));
            this.pictureBoxLoader.InitialImage = null;
            this.pictureBoxLoader.Location = new System.Drawing.Point(75, 205);
            this.pictureBoxLoader.Margin = new System.Windows.Forms.Padding(70, 200, 50, 50);
            this.pictureBoxLoader.Name = "pictureBoxLoader";
            this.pictureBoxLoader.Size = new System.Drawing.Size(201, 202);
            this.pictureBoxLoader.TabIndex = 0;
            this.pictureBoxLoader.TabStop = false;
            this.pictureBoxLoader.Visible = false;
            this.pictureBoxLoader.Click += new System.EventHandler(this.pictureBoxLoader_click);
            // 
            // buttonGenerate
            // 
            this.buttonGenerate.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(200)))), ((int)(((byte)(60)))), ((int)(((byte)(28)))));
            this.buttonGenerate.FlatAppearance.BorderSize = 0;
            this.buttonGenerate.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonGenerate.Font = new System.Drawing.Font("Arial", 10F, System.Drawing.FontStyle.Bold);
            this.buttonGenerate.ForeColor = System.Drawing.Color.White;
            this.buttonGenerate.Location = new System.Drawing.Point(110, 14);
            this.buttonGenerate.Margin = new System.Windows.Forms.Padding(4);
            this.buttonGenerate.Name = "buttonGenerate";
            this.buttonGenerate.Size = new System.Drawing.Size(135, 30);
            this.buttonGenerate.TabIndex = 2;
            this.buttonGenerate.Text = " Generate ";
            this.buttonGenerate.UseVisualStyleBackColor = false;
            this.buttonGenerate.Click += new System.EventHandler(this.myButton_Click);
            this.buttonGenerate.MouseEnter += new System.EventHandler(this.buttonGenerate_MouseEnter);
            this.buttonGenerate.MouseLeave += new System.EventHandler(this.buttonGenerate_MouseLeave);
            // 
            // buttonLoadMore
            // 
            this.buttonLoadMore.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(200)))), ((int)(((byte)(60)))), ((int)(((byte)(28)))));
            this.buttonLoadMore.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonLoadMore.Font = new System.Drawing.Font("Arial", 10F, System.Drawing.FontStyle.Bold);
            this.buttonLoadMore.ForeColor = System.Drawing.Color.White;
            this.buttonLoadMore.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonLoadMore.Location = new System.Drawing.Point(67, 668);
            this.buttonLoadMore.Margin = new System.Windows.Forms.Padding(4);
            this.buttonLoadMore.Name = "buttonLoadMore";
            this.buttonLoadMore.Size = new System.Drawing.Size(280, 46);
            this.buttonLoadMore.TabIndex = 3;
            this.buttonLoadMore.Text = " See more Design Ideas";
            this.buttonLoadMore.UseVisualStyleBackColor = false;
            this.buttonLoadMore.Visible = false;
            this.buttonLoadMore.Click += new System.EventHandler(this.loadMore);
            this.buttonLoadMore.MouseEnter += new System.EventHandler(this.buttonLoadMore_MouseEnter);
            this.buttonLoadMore.MouseLeave += new System.EventHandler(this.buttonLoadMore_MouseLeave);
            // 
            // UserControl1
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Inherit;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(248)))), ((int)(((byte)(244)))), ((int)(((byte)(244)))));
            this.Controls.Add(this.buttonGenerate);
            this.Controls.Add(this.buttonLoadMore);
            this.Controls.Add(this.flowLayoutPanelSlides);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "UserControl1";
            this.Size = new System.Drawing.Size(408, 718);
            this.Load += new System.EventHandler(this.UserControl1_Load);
            this.flowLayoutPanelSlides.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxLoader)).EndInit();
            this.ResumeLayout(false);

        }        private void buttonGenerate_MouseEnter(object sender, EventArgs e)        {            buttonGenerate.BackColor = System.Drawing.Color.FromArgb(112, 36, 12); // #70240c
        }

        // Event handler for buttonGenerate hover leave
        private void buttonGenerate_MouseLeave(object sender, EventArgs e)        {            buttonGenerate.BackColor = System.Drawing.Color.FromArgb(200, 60, 28); // Original color
        }

        // Event handler for buttonLoadMore hover enter
        private void buttonLoadMore_MouseEnter(object sender, EventArgs e)        {            buttonLoadMore.BackColor = System.Drawing.Color.FromArgb(112, 36, 12); // #70240c
        }

        // Event handler for buttonLoadMore hover leave
        private void buttonLoadMore_MouseLeave(object sender, EventArgs e)        {            buttonLoadMore.BackColor = System.Drawing.Color.FromArgb(200, 60, 28); // Original color
        }




        #endregion
        // Event handler for Load More button

    }}