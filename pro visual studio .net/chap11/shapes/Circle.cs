using System;

namespace Apress.ProVisualStudio.chap11.shapes
{
	/// <summary>
	/// Summary description for Circle.
	/// </summary>
	public class Circle : Shape
	{
		public float radius;   // radius of circle

		public Circle()
		{
			// constructor
			// logic 
			// here
			x = y = 0; radius = 0;
		}
		public Circle(int x, int y, float radius)
		{
			this.x = x;
			this.y = y;
			this.radius = radius;
		}

		public override float Area()
		{
			return (float)(System.Math.PI) * (radius * radius);  // pie are squared :)
		}
		public float radiusLength 
		{
			get 
			{
				return radius;
			} 
			set 
			{
				this.radius = value;
			}
		}

	}
}
