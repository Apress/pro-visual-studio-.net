using System;

namespace Apress.ProVisualStudio.chap11.shapes
{
	public class Triangle : Shape
	{
		private float Base;		// base (no initial cap) is a keyword
		private float Height;

		public Triangle()
		{
			x = y = 0;
			Base = Height = 0;
		}
		public Triangle(int x, int y, float Base, float Height)
		{
			this.x = x;
			this.y = y;
			this.Base = Base;
			this.Height = Height;
		}

		public override float Area()
		{
			return (float)0.5 * Base * Height;  // triangle area = 1/2 base * height
		}
		public float baseLength 
		{
			get 
			{
				return Base;
			} 
			set 
			{
				this.Base = value;
			}
		}
		public float heightLength 
		{
			get 
			{
				return Height;
			} 
			set 
			{
				this.Height = value;
			}
		}

	}

}
