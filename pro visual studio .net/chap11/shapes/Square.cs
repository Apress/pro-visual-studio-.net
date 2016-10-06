using System;

namespace Apress.ProVisualStudio.chap11.shapes
{
	public class Square : Shape
	{
		private float side;		// base (no initial cap) is a keyword

		public Square()
		{
			x = y = 0;
			side = 0;
		}
		public Square(int x, int y, float side)
		{
			this.x = x;
			this.y = y;
			this.side = side;
		}

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
