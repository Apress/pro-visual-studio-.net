using System;


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
		public Shape(int x, int y)
		{
			this.x = x;
			this.y = y;
		}

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


