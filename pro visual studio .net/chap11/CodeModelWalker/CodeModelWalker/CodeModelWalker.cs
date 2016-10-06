using System;
using System.Diagnostics;
using EnvDTE;

namespace Apress.ProVisualStudio.chap11.CodeModelWalker
{
	/// <summary>
	/// Summary description for CodeModelWalker.
	/// </summary>
	public class CodeModelWalker
	{
		public CodeModelWalker()
		{
		}

		public static void WalkDTE(EnvDTE._DTE applicationObject)
		{
			do_DTE(applicationObject);
			WalkSolution(applicationObject.Solution);
		}
		public static void WalkSolution(EnvDTE.Solution solution)
		{
			doSolution(solution);
			WalkProjects(solution.Projects);
		}

		public static void WalkProjects(EnvDTE.Projects projects)
		{
			foreach (EnvDTE.Project proj in projects)
			{
				doProject(proj);
				foreach (EnvDTE.ProjectItem projItem in proj.ProjectItems)
				{
					doProjectItem(projItem);
					WalkFileCodeModel(projItem.FileCodeModel, "");
				}
				WalkCodeModel(proj.CodeModel, "");
			}
		}
		public static void WalkFileCodeModel(EnvDTE.FileCodeModel fileCodeModel, string indent) 
		{
			doFileCodeModel(fileCodeModel);
			WalkCodeElements(fileCodeModel.CodeElements, indent);
		}
		public static void WalkCodeModel(EnvDTE.CodeModel codeModel, string indent) 
		{
			doCodeModel(codeModel);
			WalkCodeElements(codeModel.CodeElements, indent);
		}

		public static void WalkCodeElements(EnvDTE.CodeElements codeElements, string indent) 
		{
			foreach (EnvDTE.CodeElement codeElem in codeElements)
			{
				if (codeElem.Kind == EnvDTE.vsCMElement.vsCMElementNamespace)
				{
					if (codeElem.Name != "System")			// walking System/Microsoft namespaces generates too much output
						if (codeElem.Name != "Microsoft")
							WalkNamespace((EnvDTE.CodeNamespace) codeElem, indent);
				} 
				else
				{
					doCodeElement(codeElem, indent);
				}
			}
		}
		public static void WalkNamespace(EnvDTE.CodeNamespace cns, string indent) 
		{
			doNamespace(cns, indent);
			foreach (EnvDTE.CodeElement codeElem in cns.Members)
			{
				switch (codeElem.Kind) 
				{
					case EnvDTE.vsCMElement.vsCMElementNamespace:
						WalkNamespace((CodeNamespace) codeElem, indent+"..."); 
						break;
					case EnvDTE.vsCMElement.vsCMElementClass:
						WalkClass((CodeClass) codeElem, indent+"...");
						break;
					case EnvDTE.vsCMElement.vsCMElementInterface:
						WalkInterface((CodeInterface) codeElem, indent+"...");
						break;
					default:
						doCodeElement(codeElem, indent);
						break;
				}
			} 
		}

		public static void WalkClass(EnvDTE.CodeClass cls, string indent) 
		{
			doClass(cls, indent);
			foreach (EnvDTE.CodeElement codeElem in cls.Bases)
			{
				doInheritsFrom(codeElem, indent);
			}
			foreach (EnvDTE.CodeElement codeElem in cls.Members)
			{
				switch (codeElem.Kind) 
				{
					case EnvDTE.vsCMElement.vsCMElementVariable:
						doCodeVariable((CodeVariable) codeElem, indent);
						break;
					case EnvDTE.vsCMElement.vsCMElementProperty:
						doCodeProperty((CodeProperty) codeElem, indent);
						break;
					case EnvDTE.vsCMElement.vsCMElementFunction:
						WalkFunction((CodeFunction) codeElem, indent+"...");
						break;
					default:
						doCodeElement(codeElem, indent);
						break;
				}
			}
		}

		public static void WalkInterface(EnvDTE.CodeInterface ifac, string indent) 
		{
			doInterface(ifac, indent);
			foreach (EnvDTE.CodeElement codeElem in ifac.Bases)
			{
				doInheritsFrom(codeElem, indent);
			}
			foreach (EnvDTE.CodeElement codeElem in ifac.Members)
			{
				switch (codeElem.Kind) 
				{
					// no EnvDTE.vsCMElement.vsCMElementVariable case for interface
					case EnvDTE.vsCMElement.vsCMElementProperty:
						doCodeProperty((CodeProperty) codeElem, indent);
						break;
					case EnvDTE.vsCMElement.vsCMElementFunction:
						WalkFunction((CodeFunction) codeElem, indent+"...");
						break;
					default:
						doCodeElement(codeElem, indent);
						break;
				}
			}

		}
		public static void WalkFunction(EnvDTE.CodeFunction func, string indent)
		{
			doFunction(func, indent);
			doParameters(func, indent);
			//doCreateXMLDocumentForFunction(func);
		}
		//------------------------ "do" methods separate from "walk" methods here
		public static void do_DTE(EnvDTE._DTE applicationObject) 
		{
			Debug.WriteLine("DTE applicationObject: {0}", applicationObject.Name);
		}
		public static void doSolution(EnvDTE.Solution solution) 
		{
			Debug.WriteLine("Solution - FullName: {0}", solution.FullName);
		}

		public static void doProject(EnvDTE.Project proj) 
		{
			Debug.WriteLine("******Project Name:"+proj.Name+"***********");
		}
		public static void doProjectItem(EnvDTE.ProjectItem projItem)
		{
			Debug.WriteLine("----- project item:"+projItem.Name+ " ------------");
		}
		public static void doFileCodeModel(EnvDTE.FileCodeModel fileCodeModel)
		{
			Debug.WriteLine("walking FileCodeModel...");
			if (fileCodeModel.CodeElements.Count <= 0) 
			{
				Debug.WriteLine("no code elements in this FileCodeModel");
			}
		}

		public static void doCodeModel(EnvDTE.CodeModel codeModel)
		{
			Debug.WriteLine("======== Walking Project.CodeModel for: "+codeModel.Parent.Name+"=========");
			Debug.WriteLine("==== code model language is: "+ codeModel.Language+" =======");
		}
		public static void doCodeElement(EnvDTE.CodeElement codeElem, string indent) 
		{
			Debug.WriteLine(indent + "code element:" + codeElem.Name +" kind: "+codeElem.Kind);
		}
		public static void doNamespace(EnvDTE.CodeNamespace cns, string indent) 
		{
			Debug.WriteLine(indent + "namespace:" + cns.Name);
		}
		public static void doClass(EnvDTE.CodeClass cls, string indent) 
		{
			Debug.WriteLine(indent + "class:" + cls.Name);
		}
		public static void doInheritsFrom(EnvDTE.CodeElement codeElem, string indent) 
		{
			Debug.WriteLine(indent + "inherits from:" + codeElem.Name);
		}
		public static void doCodeVariable(EnvDTE.CodeVariable var, string indent)
		{
			Debug.WriteLine(indent+"..."+"variable: "+var.Name + " type: " + var.Type.AsString);
		}
		public static void doCodeProperty(EnvDTE.CodeProperty prop, string indent) 
		{
			Debug.WriteLine(indent+"..."+"property: "+ prop.Name + " type: " + prop.Type.AsString);
		}
		public static void doInterface(EnvDTE.CodeInterface ifac, string indent) 
		{
			Debug.WriteLine(indent + "interface:" + ifac.Name);
		}
		public static void doFunction(EnvDTE.CodeFunction func, string indent) 
		{
			Debug.WriteLine(indent+"function: "+ func.Name + "() returns: "+func.Type.AsString);
			Debug.WriteLine(indent+"kind: "+ func.FunctionKind.ToString());
		}
		public static void doParameters(EnvDTE.CodeFunction func, string indent) 
		{
			if (func.Parameters.Count <= 0) 
			{
				Debug.WriteLine(indent+"..."+"no parameters");
			} 
			else 
			{
				foreach (CodeParameter param in func.Parameters) 
				{
					Debug.WriteLine(indent+"..."+"parameter: "+ param.Name + " type: " + param.Type.AsString);
				}
			}
		}
		public static void doCreateXMLDocumentForFunction(EnvDTE.CodeFunction func)
		{
			// probably should use StringBuilder here for better performance/memory use
			// but am not as readability is main purpose of example 
			string xmlComments = "";

			xmlComments +=
				"\t\t/// <summary> " + func.Name+ "()";

			switch (func.FunctionKind) 
			{
				case vsCMFunction.vsCMFunctionConstructor:
					xmlComments += " constructor"; break;
				case vsCMFunction.vsCMFunctionDestructor:
					xmlComments += " destructor"; break;
				default: break;
			}

			if (func.Parameters.Count <= 0) 
			{
				xmlComments += ".  No parameters";
			} 
			xmlComments += ". </summary> \n";

			foreach (CodeParameter param in func.Parameters) 
			{
				xmlComments += 
					"\t\t/// <param name=\""+ param.Name + "\"> type: " +
					param.Type.AsString + "</param> \n";
			}
			xmlComments += 			
				"\t\t/// <returns> "+func.Type.AsString + "</returns>\n";


			Debug.WriteLine(xmlComments);

			if (func.DocComment.Length <= 1)
			{
				// this does not work: func.DocComment = xmlComments;
				// editing does work
				Debug.WriteLine("writing XML comments to source file...");
				EnvDTE.TextPoint tp = func.GetStartPoint(EnvDTE.vsCMPart.vsCMPartWholeWithAttributes);
				EnvDTE.EditPoint ep = tp.CreateEditPoint();
				ep.StartOfLine();
				ep.Insert(xmlComments);
			} 
			else 
			{
				Debug.WriteLine("XML comments already present in source file:");
				Debug.WriteLine(func.DocComment);
				Debug.WriteLine("-- end doc comment --");
			}

		}
	}
}
