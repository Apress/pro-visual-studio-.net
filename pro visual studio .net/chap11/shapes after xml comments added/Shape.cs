using System;

// comment
// here

namespace Apress.ProVisualStudio.chap11.shapes
{
	/// <summary>
	/// Summary description for Class1.
	/// </summary>
	public abstract class Shape
	{
		public int x;  // x cartesian coordinate
		public int y;  // y cartesian coordinate

		public Shape()
		{
			// Constructor
			// Logic
			// Here	
			x = 0;
			y = 0;
		}
		/// <summary> Shape() constructor. </summary> 
		/// <param name="x"> type: int</param> 
		/// <param name="y"> type: int</param> 
		/// <returns> void</returns>
		public Shape(int x, int y)
		{
			this.x = x;
			this.y = y;
		}

		/// <summary> Area().  No parameters. </summary> 
		/// <returns> float</returns>
		public abstract float Area();

		public int xCoordinate 
		{
			get 
			{
				return this.x;
			} set 
			  {
				  this.x = value;
			  }
		}
	
		public int yCoordinate 
		{
			get 
			{
				return this.y;
			} 
			set 
			{
				  this.y = value;
			}
		}
	}
}


