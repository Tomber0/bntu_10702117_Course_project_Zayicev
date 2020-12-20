using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
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
        static int plane = 0;
        static int curve = 0;
        static int stBox = 0;
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        List<GeometryPoint> geometryPoints = new List<GeometryPoint>();
        int Count { get; set; }
        private void Form1_Load(object sender, EventArgs e)
        {

            Settings = new Settings() { CubeX = 1000.0f, CubeY = 1000.0f, CubeZ = 1000.0f, BassR = 1000.0f, CutR = 1000.0f };
            //Заполнение формы
            dataGridView1.ColumnCount = 3;
            dataGridView1.RowCount = 2;
            dataGridView1.RowHeadersWidth = 50;
            for (int i = 0; i < dataGridView1.RowCount; i++)
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
            for (int i = 0; i < geometryPoints.Count; i++)
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
            Count = dataGridView1.RowCount-1;
            for (int i = 0; i < Count; i++)
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
        private void BuidPlane()
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
        
        //построить кривую
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


            curve++;
            swModel.InsertCurveFileBegin();// начало проведения ломаной
            //var status = swModel.InsertCurveFilePoint(0, 0, 0);

            for (int i = 0; i < geometryPoints.Count; i++)
            {
                swModel.InsertCurveFilePoint(Convert.ToInt32(geometryPoints.ToArray()[i].X), Convert.ToInt32(geometryPoints.ToArray()[i].Y), Convert.ToInt32(geometryPoints.ToArray()[i].Z));
            }
            swModel.InsertCurveFileEnd();// конец проведения ломаной
            var sManager = swModel.SelectionManager;
            var selObj = sManager.GetSelectedObject5(1);
            selObj.name = $"Curve991" + curve.ToString();//test
            var boolStatus = swModel.Extension.SelectByID2("", "POINTREF", Convert.ToInt32(geometryPoints.ToArray()[0].X), Convert.ToInt32(geometryPoints.ToArray()[0].Y), Convert.ToInt32(geometryPoints.ToArray()[0].Z), true, 1, null, 0);

            var swRefPlane = (RefPlane)swFeatureManager.InsertRefPlane(2, 0, 4, 0, 0, 0);

            //var bResult = swModel.Extension.SelectByID2("", "PLANE", geometryPoints.ToArray()[0].X, geometryPoints.ToArray()[0].Y, geometryPoints.ToArray()[0].Z, false, 0, null, 0);
            // status = swModelDocExt.SelectByID2("Curve1", "REFERENCECURVES", 0, 0,0, false, 0, null, 0);
            swModel.ClearSelection2(true);


        }




        /*
         * 
         * 
         * ?/
        
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

            var swRefPlane = (RefPlane)swFeatureManager.InsertRefPlane(4, 0, 2, 0, 0, 0);




        }
         
          
         */

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


        // создание модели по траектории
        private void button5_Click(object sender, EventArgs e)
        {

            if (geometryPoints.Count == 0)
            {
                MessageBox.Show("Сначала постройте плоскость!");
                return;
            }

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


            var bResult = true;
            bResult = swModel.Extension.SelectByID2("", "PLANE", geometryPoints.ToArray()[0].X, geometryPoints.ToArray()[0].Y, geometryPoints.ToArray()[0].Z, false, 0, null, 0);
            swModel.SketchManager.InsertSketch(true);
            swModel.ClearSelection2(true);



            var circle = swModel.SketchManager.CreateCircleByRadius(0, 0, 0, Settings.BassR/1000);
            var activeSketch = swModel.GetActiveSketch2();
            activeSketch.Name = $"Sketch991" + plane.ToString();

            swModel.SketchManager.InsertSketch(true);
            swModel.ClearSelection2(true);

            bResult = swModel.Extension.SelectByID2($"Sketch991" + plane.ToString(), "SKETCH", 0, 0, 0, false, 1, null, 0);
            //test//bResult = swModel.Extension.SelectByID2("", "REFERENCECURVES", geometryPoints.ToArray()[0].X, geometryPoints.ToArray()[0].Y, geometryPoints.ToArray()[0].Z,true,2,null,0);
            bResult = swModel.Extension.SelectByID2($"Curve991" + curve.ToString(), "REFERENCECURVES", 0, 0, 0, true, 2, null, 0);//test
            var swFeature = (Feature)swModel.FeatureManager.InsertProtrusionSwept4(false, false, 0, false, false, 0, 0, false, 0, 0,
0, 0, true, true, true, 0, true, false, 0.1, 0);

            plane++;
            //var mf = swModel.FeatureManager.InsertProtrusionSwept4(false, false, 0, false, false, 0, 0, false, 0, 0, 0, 0, true, true, true, 0, true, false, 0.01,1);
            //var swFeature = (Feature)swModel.FeatureManager.InsertProtrusionSwept4(false, false, (int)swTwistControlType_e.swTwistControlFollowPath, false, false, (int)swTangencyType_e.swTangencyNone, (int)swTangencyType_e.swTangencyNone, false, 0, 0, (int)swThinWallType_e.swThinWallOneDirection, 0, true, true, true, 0, true, false, 0.1, 0);

        }


        //вырез через кривую
        private void button6_Click(object sender, EventArgs e)
        {
            if (geometryPoints.Count == 0)
            {
                MessageBox.Show("Сначала постройте плоскость!");
                return;
            }

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


            var bResult = true;
            bResult = swModel.Extension.SelectByID2("", "PLANE", geometryPoints.ToArray()[0].X, geometryPoints.ToArray()[0].Y, geometryPoints.ToArray()[0].Z, false, 0, null, 0);
            swModel.SketchManager.InsertSketch(true);
            swModel.ClearSelection2(true);



            var circle = swModel.SketchManager.CreateCircleByRadius(0, 0, 0, Settings.CutR/1000);
            var activeSketch = swModel.GetActiveSketch2();
            activeSketch.Name = $"Sketch991" + plane.ToString();

            swModel.SketchManager.InsertSketch(true);
            swModel.ClearSelection2(true);

            bResult = swModel.Extension.SelectByID2($"Sketch991" + plane.ToString(), "SKETCH", 0, 0, 0, false, 1, null, 0);
            //test
            bResult = swModel.Extension.SelectByID2($"Curve991" + curve.ToString(), "REFERENCECURVES", 0, 0, 0, true, 4, null, 0);
            //
            //test //bResult = swModel.Extension.SelectByID2("", "REFERENCECURVES", geometryPoints.ToArray()[0].X, geometryPoints.ToArray()[0].Y, geometryPoints.ToArray()[0].Z, true, 4, null, 0);
            /*            var swFeature = (Feature)swModel.FeatureManager.InsertProtrusionSwept4(false, false, 0, false, false, 0, 0, false, 0, 0,
            0, 0, true, true, true, 0, true, false, 0.1, 0);
            */

            var swFeature2 = (Feature)swModel.FeatureManager.InsertCutSwept4(false, true, 0, false, false, 0, 0, false, 0, 0, 0, 0, true, true, 0, true, true, true, false);
            //var swFeature = (Feature)swModel.FeatureManager.InsertCutSwept5(false, true, 0, false, false, 0, 0, false, 0, 0, 0, 0, true, true, 0, true, true, true, false, true, 0.1, 0);

            plane++;


            /*            var SwApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application"); //проверка на то, включен ли солид

                        IModelDoc2 swModel = SwApp.IActiveDoc2;     //стандартное присвоение

                        var swFeature = (Feature)swModel.FeatureManager.InsertProtrusionSwept4(false, false, 0, false, false, 0, 0, false, 0, 0,
            0, 0, true, true, true, 0, true, false, 0.1, 0);
            */
        }

        //Построение параллелепипеда
        private void button7_Click(object sender, EventArgs e)
        {

            if (geometryPoints.Count == 0) 
            {
                MessageBox.Show("Сначала постройте плоскость!");
                return;
            }
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


            var bResult = true;
            bResult = swModel.Extension.SelectByID2("", "PLANE", geometryPoints.ToArray()[0].X, geometryPoints.ToArray()[0].Y, geometryPoints.ToArray()[0].Z, false, 0, null, 0);
            swModel.SketchManager.InsertSketch(true);
            swModel.ClearSelection2(true);



            var box = swModel.SketchManager.CreateCenterRectangle(0, 0, 0, Settings.CubeY/2000, Settings.CubeX/2000, Settings.CubeZ/1000);

            swModel.FeatureManager.FeatureExtrusion2(true, false, false, 0, 0, Settings.CubeZ/1000, Settings.CubeZ/1000, false, false, false, false, 1.74532925199433E-02, 1.74532925199433E-02, false, false, false, false, true, true, true, 0, 0, false);
            swModel.ClearSelection2(true);
            //swModel.SelectionManager.EnableContourSelection =false;
            //var activeSketch = swModel.GetActiveSketch2();
            //activeSketch.Name = $"SketcBoxh991" + box.ToString();



            stBox++;
        }

        private void Form1_KeyPress(object sender, KeyPressEventArgs e)
        {


        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {

        }

        private void dataGridView1_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            dataGridView1.Rows[e.RowIndex].HeaderCell.Value = $"{e.RowIndex+1}";

        }

        private void dataGridView1_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            dataGridView1.Rows[e.RowIndex].HeaderCell.Value = $"{e.RowIndex+1}";

        }
        SizeForm sizeformRef;
        public Settings Settings;

        private void button1_Click_1(object sender, EventArgs e)
        {
            sizeformRef = new SizeForm();
            sizeformRef.settings = Settings;
            sizeformRef.form1Ref = this;
            sizeformRef.ShowDialog();
        }

        public void UpdateSettings(Settings newSettings) 
        {
            Settings = newSettings;
        }
    }
}

/*            swModel.InsertProtrusionSwept4(false, false, (int)swTwistControlType_e.swTwistControlFollowPathFirstGuideCurve, false, false, (int)swTangencyType_e.swTangencyNone, (int)swTangencyType_e.swTangencyNone, true, 0, 0, (int)swThinWallType_e.swThinWallMidPlane);


            //var mf2 = (Feature)swModel.InsertProtrusionSwept4(false, false, (int)swTwistControlType_e.swTwistControlFollowPath, false, false, (int)swTangencyType_e.swTangencyNone, (int)swTangencyType_e.swTangencyNone, false, 0, 0, (int)swThinWallType_e.swThinWallOneDirection, 0, true, true, true, 0, true, false, 0, 0);

*/            //var mf2 = (Feature)swModel.FeatureManager.InsertProtrusionSwept4(false, false, (int)swTwistControlType_e.swTwistControlFollowPathFirstGuideCurve, false, (int)swTangencyType_e.swTangencyNone, (int)swTangencyType_e.swTangencyNone, false, 0, 0, (int)swThinWallType_e.swThinWallOneDirection, 0, true, true, true, 0, true, false, 0, 0);