using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;


namespace itConference
{
    public partial class formCam : Form
    {
        private ClassVideo pcc;

        public formCam()
        {
            InitializeComponent();
            //pcc = new ClassVideo(this.pictureBox1.Handle, 0, 0, 160, 120);  //panelCamera is Control panel
            pcc = new ClassVideo(this.pictureBox1.Handle, 0, 0, 218, 302);
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            pcc.Start();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            pcc.Stop();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (File.Exists(Directory.GetCurrentDirectory() + "\\imgtemp.jpg")) //存在文件
            {
                File.Delete(Directory.GetCurrentDirectory() + "\\imgtemp.jpg");
            }
            if (File.Exists(Directory.GetCurrentDirectory() + "\\imgtemp1.jpg")) //存在文件
            {
                File.Delete(Directory.GetCurrentDirectory() + "\\imgtemp1.jpg");
            }
            pcc.GrabImage(Directory.GetCurrentDirectory() + "\\imgtemp.jpg");
            

            FileStream photoStream = new FileStream(Directory.GetCurrentDirectory() + "\\imgtemp.jpg", FileMode.Open, FileAccess.Read);
            byte[] photoBytes = new byte[photoStream.Length];
            photoStream.Read(photoBytes, 0, (int)photoStream.Length);
            photoStream.Close();

            MemoryStream StreamPhoto = new MemoryStream(photoBytes);
            //this.pictureBoxGet.Image = Image.FromStream(StreamPhoto);
            Bitmap bitmap1 = (Bitmap)Bitmap.FromFile(Directory.GetCurrentDirectory() + "\\imgtemp.jpg");
            bitmap1.RotateFlip(RotateFlipType.Rotate90FlipY);

            this.pictureBoxGet.Image = bitmap1;
            bitmap1.Save(Directory.GetCurrentDirectory() + "\\imgtemp1.jpg");

            File.Delete(Directory.GetCurrentDirectory() + "\\imgtemp.jpg");
            File.Delete(Directory.GetCurrentDirectory() + "\\imgtemp1.jpg");

        }

        private void formCam_Load(object sender, EventArgs e)
        {
            pcc.Start();
        }

        private void btnYES_Click(object sender, EventArgs e)
        {

        }
    }
}
