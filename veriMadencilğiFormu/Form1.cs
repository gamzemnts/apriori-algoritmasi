using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;


namespace veriMadencilğiFormu
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        int[] destek;
        string[] str;
        int[,] dizi;
        string path;
        Workbook wb;
        Worksheet ws;
        int satir = 0,sutun = 0;


        private void button1_Click(object sender, EventArgs e)
        {
            sil();
            dosya();
            olustur();
            dizi = new int[sutun,satir];
            destek = new int[sutun];
            str = new string[sutun];
            
            bilgi();
            tekDestek();
            siralama();
            kural1();
        }

        private void sil()
        {
            listBox1.Items.Clear();
            listBox2.Items.Clear();
            listBox3.Items.Clear();
        }

        private void dosya()
        {
            try
            {
                openFileDialog1.ShowDialog();
                path = openFileDialog1.FileName;

            }
            catch
            {
            }
        }

        void bagla(){
            try
            {
                _Application excel = new _Excel.Application();
                wb = excel.Workbooks.Open(path);
                ws = wb.Worksheets[1];
            }
            catch 
            { 
                
            }
        }

        private void olustur()
        {
            bagla();
            try
            {
                while (true)
                {

                    if (ws.Cells[satir + 2, 1].Value2 != null)
                    {
                        satir++;
                    }
                    else
                    {
                        break;

                    }

                }



                while (true)
                {

                    if (ws.Cells[1,sutun + 1 ].Value2 != null)
                    {
                        sutun++;
                    }
                    else
                    {
                        break;

                    }
                }
            }
            finally
            {
                wb.Close(path);
            }


           
        }

        private void bilgi()
        {
            bagla();
            try
            {
                for (int i = 0; i < sutun; i++)
                {
                    for (int j = 0; j <satir ; j++)
                    {
                       dizi[i, j] =Convert.ToInt32(ws.Cells[j + 2, i + 1].Value2);
                    }
                }
                for (int i = 0; i < sutun; i++)
                {
                    str[i] = Convert.ToString(ws.Cells[1, i + 1].Value2);
                }

                
            }
            finally
            {
                wb.Close(path);
            }
            
        }

        private void tekDestek()
        {
            for (int i = 0; i < sutun; i++)
            {
                for (int j = 0; j < satir; j++)
                {
                    if ( dizi[i, j]== 1)
                    {
                        destek[i] ++;
                    }

                }
                listBox3.Items.Add(str[i]+": "+destek[i]);
            }
            
        }

        private void siralama()
        {
            for (int i = 0; i < sutun; i++)
            {
                for (int j = 0; j < sutun; j++)
                {
                    if (destek[i] > destek[j])
                    {
                        int yedek = destek[i];
                        destek[i] = destek[j];
                        destek[j] = yedek;

                        string yedek2 = str[i];
                        str[i] = str[j];
                        str[j] = yedek2;

                        for (int k = 0; k < satir; k++)
                        { 
                            int yedek3 = dizi[i, k];
                            dizi[i, k] = dizi[j, k];
                            dizi[j, k] = yedek3;
                        }
                    }
                }
            }
        }

        private void kural1()
        {
           

            for (int i = 0; i < sutun; i++)
            {
                if (destek[i] >= satir/2)
                {
                    for (int j = i+1; j < sutun; j++)
                    {
                        if (destek[j] >= satir / 2)
                        {
                            double destek1 = 0, guven = 0;
                            for (int k = 0; k < satir; k++)
                            {
                                if (dizi[i, k] == 1 && dizi[j, k] == 1)
                                {
                                    destek1++;
                                    guven++;
                                }
                            }
                            double x = destek1;
                            destek1 = destek1 *100/ satir;
                            guven = guven*100/ destek[i];
                            if (destek1 >= 50 && guven >= 50)
                            {
                                listBox1.Items.Add(str[i]+":"+str[j]+":"+guven+"%");
                                kural2(i, j, x);
                            }
                        }
                    }
                }
            }
        }

        private void kural2(int i, int j, double x)
        {


            for (int k = j + 1; k <sutun  ; k++)
            {

                if (destek[k] >= satir/2)
                {
                    double destek1 = 0, guven = 0;
                    for (int l = 0; l < satir; l++)
                    {
                        if (dizi[i, l] == 1 && dizi[j, l] == 1 && dizi[k, l] == 1)
                        {
                            destek1++;
                            guven++;
                        }
                    }
                    destek1 = destek1*100 / satir;
                    guven = guven*100/ x;
                    if (destek1 >= 50 && guven >= 50)
                    {
                        listBox2.Items.Add(str[i] + ":" + str[j] + ":" + str[k] + ":" + guven + "%");
                    }
                }
            }
        }
    }
}
