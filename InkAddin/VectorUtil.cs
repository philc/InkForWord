using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;

namespace InkAddin
{
    /// <summary>
    /// Some vector utilities needed for rotation/distance calculation.
    /// </summary>
    class Vector
    {
        public double X = 0;
        public double Y = 0;

        // An angle is from 0 to 180 degrees, starting clockwise from the X-Axis;
        // if it's counter clockwise from the X-axis, it's 0 to -180 degrees.
        public double Angle = 0;

        private static Vector XAxis = new Vector(1, 0);
        public Vector(double x, double y)
        {
            this.X = x;
            this.Y = y;

            // Case where we're the X axis. Don't need to compuete angle.
            if (this.X == 1 && this.Y == 0)
                this.Angle = 0;
            else
                this.Angle = AngleBetween(XAxis);

            if (Y < 0)
                this.Angle = -this.Angle;
        }
        public Vector(Point p, Point origin)
            : this(p.X - origin.X, p.Y - origin.Y)
        {
        }
        public double AngleBetween(Vector v2)
        {
            // u . v = ||u|| ||v|| cos(theta)

            double cosTheta = this.DotProduct(v2) / (Length(this) * Length(v2));
            // Expects radians
            double angle = Math.Acos(cosTheta);
            // Convert to degrees
            return (angle * (180 / Math.PI));
        }
        private static double Length(Vector vector)
        {
            return (float)Math.Sqrt(vector.DotProduct(vector));
        }
        private double DotProduct(Vector q)
        {
            return this.X * q.X + this.Y * q.Y;
        }
        public Vector MakeUnitVector()
        {
         // Length will never be 0 in this method; vector can't be orthagonal to itself.
         double length = Length(this);
         return new Vector(this.X/ length, this.Y / length);
        }
    }
}
