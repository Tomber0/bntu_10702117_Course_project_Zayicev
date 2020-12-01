using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace bntu_10702117_Course_project_Zayicev
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        List<GeometryPoint> geometryPoints = new List<GeometryPoint>();
        private void Form1_Load(object sender, EventArgs e)
        {

            //Заполнение формы
            dataGridView1.ColumnCount = 3;
            dataGridView1.RowCount = 3;
            dataGridView1.RowHeadersWidth = 50;
            for (int i = 0; i < 3; i++)
            {
                dataGridView1.Rows[i].HeaderCell.Value = $"{i + 1}";
            }
            //dataGridView1.Rows.Add($"X",$"Y",$"Z");
        }
        private void button1_Click(object sender, EventArgs e)
        {
            SldWorks SwApp;

            try
            {
                SwApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application"); //проверка на то, включен ли солид
            }
            catch
            {
                MessageBox.Show("Пожалуйста, включите SolidWorks."); //если не влкючен
                return;
            }
            FillPoints();
            IModelDoc2 swModel = SwApp.IActiveDoc2;     //стандартное присвоение
            swModel.SketchManager.InsertSketch(true); //создать скетч
            //Создать 3 точки
            for (int i = 0; i < 3; i++)
            {
                /*                int x = Convert.ToInt32(geometryPoints.ToArray()[i].X);
                                int y = Convert.ToInt32(geometryPoints.ToArray()[i].Y);
                                int z = Convert.ToInt32(geometryPoints.ToArray()[i].Z);
                */
                swModel.CreatePoint(Convert.ToInt32(geometryPoints.ToArray()[i].X), Convert.ToInt32(geometryPoints.ToArray()[i].Y), Convert.ToInt32(geometryPoints.ToArray()[i].Z));
            }
            swModel.ClearSelection2(true);
            swModel.SketchManager.InsertSketch(true);


        }
        private bool ClearPoints() 
        {

            geometryPoints.Clear();
            return true;
        }
        private bool FillPoints() 
        {
            ClearPoints();
            for (int i = 0; i < 3; i++)
            {
                double[] cellArray = new double[3];
                for (int j = 0; j < 3; j++)
                {
                    try
                    {
                        cellArray[j] = Convert.ToDouble(dataGridView1.Rows[i].Cells[j].Value);

                    }
                    catch (Exception)
                    {
                        MessageBox.Show($"Проверьте ячейки в {i} строке, они должы содержать числовые данные");
                        throw;
                    }
                }
                geometryPoints.Add(new GeometryPoint(cellArray));
            }
            return true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SldWorks SwApp;

            try
            {
                SwApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application"); //проверка на то, включен ли солид
            }
            catch
            {
                MessageBox.Show("Пожалуйста, включите SolidWorks."); //если не влкючен
                return;
            }
            FillPoints();

            IModelDoc2 swModel = SwApp.IActiveDoc2;
            FeatureManager swFeatureManager = default(FeatureManager);

            swFeatureManager = (FeatureManager)swModel.FeatureManager;
            var swModelDocExt = (ModelDocExtension)swModel.Extension;



            swModel.InsertCurveFileBegin();// начало проведения ломаной
            //var status = swModel.InsertCurveFilePoint(0, 0, 0);

            for (int i = 0; i < 3; i++)
            {
                swModel.InsertCurveFilePoint(Convert.ToInt32(geometryPoints.ToArray()[i].X), Convert.ToInt32(geometryPoints.ToArray()[i].Y), Convert.ToInt32(geometryPoints.ToArray()[i].Z));
            }
            swModel.InsertCurveFileEnd();// конец проведения ломаной
            var boolStatus =             swModel.Extension.SelectByID2("", "POINTREF", Convert.ToInt32(geometryPoints.ToArray()[0].X), Convert.ToInt32(geometryPoints.ToArray()[0].Y), Convert.ToInt32(geometryPoints.ToArray()[0].Z), true ,1, null, 0);

            var swRefPlane = (RefPlane)swFeatureManager.InsertRefPlane(2, 0, 4, 0, 0, 0);




            // status = swModelDocExt.SelectByID2("Curve1", "REFERENCECURVES", 0, 0,0, false, 0, null, 0);


        }

        private void button3_Click(object sender, EventArgs e)
        {
            SldWorks SwApp;

            try
            {
                SwApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application"); //проверка на то, включен ли солид
            }
            catch
            {
                MessageBox.Show("Пожалуйста, включите SolidWorks."); //если не влкючен
                return;
            }

            IModelDoc2 swModel = SwApp.IActiveDoc2;     //стандартное присвоение

            FeatureManager swFeatureManager = default(FeatureManager);



            swFeatureManager = (FeatureManager)swModel.FeatureManager;
            var swModelDocExt = (ModelDocExtension)swModel.Extension;



            //var boolstatus = swModelDocExt.SelectByID2("", "SKETCHPOINT", 0.028424218552, 0.07057725774359, 0, true, 0, null, 0);
            

            swModel.ClearSelection2(true);
            swModel.SketchManager.InsertSketch(true);

            var bools = new List<bool>();
            foreach (var item in geometryPoints)
            {
                var boolstatus = swModelDocExt.SelectByID2("", "REFERENCECURVES", Convert.ToInt32(item.X), Convert.ToInt32(item.Y), Convert.ToInt32(item.Z), false, 0, null, 0);
                bools.Add(boolstatus);
            }
            var boolStatus2 = swModel.Extension.SelectByID2("Point1", "SKETCHPOINT", Convert.ToInt32(geometryPoints.ToArray()[0].X), Convert.ToInt32(geometryPoints.ToArray()[0].Y), Convert.ToInt32(geometryPoints.ToArray()[0].Z), false, 0, null, 0);

            var swRefPlane = (RefPlane)swFeatureManager.InsertRefPlane(4,0, 2, 0, 0, 0);




        }

        private void button4_Click(object sender, EventArgs e)
        {
            var SwApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application"); //проверка на то, включен ли солид

            IModelDoc2 swModel = SwApp.IActiveDoc2;


            swModel.InsertCurveFileBegin();
            var status = swModel.InsertCurveFilePoint(0, 0, 0);
            status = swModel.InsertCurveFilePoint(0, 0, 0.0127);
            status = swModel.InsertCurveFilePoint(0, 0, 0.0254);
            status = swModel.InsertCurveFilePoint(0, 0, 0.0381);
            status = swModel.InsertCurveFilePoint(0, 0.0254, 0.0381);
            status = swModel.InsertCurveFilePoint(0, 0.0381, 0.0381);
            status = swModel.InsertCurveFileEnd();

            ModelDocExtension swModelDocExt = default(ModelDocExtension);
            SelectionMgr swSelectionMgr = default(SelectionMgr);

            //Get free point curve feature
            swModelDocExt = (ModelDocExtension)swModel.Extension;
            swSelectionMgr = (SelectionMgr)swModel.SelectionManager;
            swModel.ClearSelection2(true);

            status = swModelDocExt.SelectByID2("Curve1", "REFERENCECURVES", 0, 0, 0, false, 0, null, 0);
           // swModel.ClearSelection2(true);

        }
    }
}
