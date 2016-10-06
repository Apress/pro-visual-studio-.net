using System;

namespace Apress.ProVisualStudio.chap11.shapes
{
	public class Square : Shape
	{
		private float side;		// base (no initial cap) is a keyword

		public Square()
		{
			// constructor
			// logic 
			// here
			x = y = 0;
			side = 0;
		}
		/// <summary> Square() constructor. </summary> 
		/// <param name="x"> type: int</param> 
		/// <param name="y"> type: int</param> 
		/// <param name="side"> type: float</param> 
		/// <returns> void</returns>
		public Square(int x, int y, float side)
		{
			this.x = x;
			this.y = y;
			this.side = side;
		}

		/// <summary> Area().  No parameters. </summary> 
		/// <returns> float</returns>
		public override float Area()
		{
			return side * side;  // square area = (length of side)^2
		}

		public float sideLength 
		{
			get 
			{
				return side;
			} 
			set 
			{
				this.side = value;
			}
		}
	}
}
