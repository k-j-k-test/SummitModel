using System;
using System.Collections.Generic;
using System.Collections;
using ModelObject;
using System.Linq;


namespace Models.Models
{
	/// <summary>
	/// Description of Model1.
	/// </summary>
	public class Model1
	{
		public Cell Cell;
		
		public Model1()
		{
			Cell = new Cell(this);
		}
		
		public void SetModel()
		{			
			Func<int, int> aa = (a) => a+1;
		}
						
		public void q1(int t)
		{
			Cell["q1", t] = 0.01;			
		}
		
		public void q2(int t)
		{
			Cell["q2", t] = 0.01 * t;	
		}
	
		public void q3(int t)
		{
			Cell["q3", t] = 0.02 * t;
		}
		
		public void lx_1(int t)
		{		
			if(t==0)
			{
				Cell["lx_1", t] = 1;				
			}
			else
			{
				Cell["lx_1", t] = (1 - Cell["q1", t-1]) * Cell["lx_1", t-1];
			}		
		}
		
		public void Dx_1(int t)
		{
			Cell["Dx_1", t] = Cell["lx_1", t] * Math.Pow(0.98, t);
		}
		
		
	}
	
}
