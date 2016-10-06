using System;

namespace Apress.ProVisualStudio.chap11.shapes
{
	public class Triangle : Shape
	{
		private float Base;		// base (no initial cap) is a keyword
		private float Height;

		/// <summary> Triangle() constructor.  No parameters. </summary> 
		/// <returns> void</returns>
		public Triangle()
		{
			// constructor
			// logic 
			// here
			x = y = 0;
			Base = Height = 0;
		}
		/// <summary> Triangle() constructor. </summary> 
		/// <param name="x"> type: int</param> 
		/// <param name="y"> type: int</param> 
		/// <param name="Base"> type: float</param> 
		/// <param name="Height"> type: float</param> 
		/// <returns> void</returns>
		public Triangle(int x, int y, float Base, float Height)
		{
			this.x = x;
			this.y = y;
			this.Base = Base;
			this.Height = Height;
		}

		/// <summary> Area().  No parameters. </summary> 
		/// <returns> float</returns>
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
