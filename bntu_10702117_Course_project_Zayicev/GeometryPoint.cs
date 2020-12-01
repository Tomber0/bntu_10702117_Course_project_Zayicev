using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace bntu_10702117_Course_project_Zayicev
{
    class GeometryPoint
    {

        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }
        public GeometryPoint(double x, double y, double z) 
        {
            X = x;
            Y = y;
            Z = z;
        }
        public GeometryPoint(double[] xyz) 
        {
            X = xyz[0];
            Y = xyz[1];
            Z = xyz[2];
        }

    }
}
