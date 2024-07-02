using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.Emit;
using Models.Models;

namespace Models
{
	class Program
	{
		public static void Main(string[] args)
		{
			RoslynCompiler.CompileAndCheckErrors(@"C:\Users\wjdrh\OneDrive\Desktop\Test\SummitModel\ModelObject\Class1.cs");
			
			Model1 m1 = new Model1();			
			
			m1.lx_1(100);
			
			Console.WriteLine(m1.Cell["lx_1",10]);
			
			SyntaxTree ss = CSharpSyntaxTree.ParseText(code);
		}
	}
	
	public class RoslynCompiler
	{
	    public static void CompileAndCheckErrors(string filePath)
	    {
	        string code = File.ReadAllText(filePath);
	
	        SyntaxTree syntaxTree = CSharpSyntaxTree.ParseText(code);
	
	        string assemblyName = Path.GetRandomFileName();
	        MetadataReference[] references = new MetadataReference[]
	        {
	            MetadataReference.CreateFromFile(typeof(object).Assembly.Location)
	            // Add references to other assemblies your code depends on
	        };
	
	        CSharpCompilation compilation = CSharpCompilation.Create(
	            assemblyName,
	            syntaxTrees: new[] { syntaxTree },
	            references: references,
	            options: new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));
	
	        using (var ms = new MemoryStream())
	        {
	            EmitResult result = compilation.Emit(ms);
	
	            if (!result.Success)
	            {
	                IEnumerable<Diagnostic> failures = result.Diagnostics.Where(diagnostic =>
	                    diagnostic.IsWarningAsError ||
	                    diagnostic.Severity == DiagnosticSeverity.Error);
	
	                foreach (Diagnostic diagnostic in failures)
	                {
	                    Console.Error.WriteLine("{0}: {1}", diagnostic.Id, diagnostic.GetMessage());
	                }
	            }
	        }
	    }
	}
}