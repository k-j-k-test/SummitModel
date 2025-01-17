﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Flee.PublicTypes;

namespace ActuLiteModel
{
    public class Model
    {
        public string Name { get; set; }
        public string CurrentSheetName { get; set; }
        public ModelEngine Engine { get; set; }
        public Dictionary<string, object> Parameters { get; set; }
        public Dictionary<string, Sheet> Sheets { get; set; }
        public Dictionary<string, CompiledCell> CompiledCells { get; set; }

        public Model(string name, ModelEngine engine)
        {
            Name = name;
            Engine = engine;
            CurrentSheetName = name;
            Parameters = new Dictionary<string, object>();
            Sheets = new Dictionary<string, Sheet>();
            CompiledCells = new Dictionary<string, CompiledCell>();
        }

        public void AddSheet(string sheetName)
        {
            if (!Sheets.ContainsKey(sheetName))
            {
                Sheets[sheetName] = new Sheet();

                foreach (var compiledCell in CompiledCells)
                {
                    Sheets[sheetName].ResisterExpression(compiledCell.Key, compiledCell.Value.Expression);
                }
            }
        }

        public void ResisterCell(string name, string formula, string description)
        {
            CompiledCell compiledCell = new CompiledCell(name, formula, description, this);
            CompiledCells[name] = compiledCell;
        }

        // 매우 큰 값의 임계점을 정의
        private const double MaxAllowedValue = 99999999999;

        public void Invoke(string cellName, int t)
        {
            try
            {
                Engine.SetModelPoint();

                AddSheet(Name);
                t = Math.Min(t, Sheet.MaxT);
                Engine.Context.Variables["t"] = t;

                if (!Sheets.TryGetValue(Name, out var sheet))
                {
                    throw new InvalidOperationException($"시트 '{Name}'을 찾을 수 없습니다.");
                }

                var expression = sheet.GetExpression(cellName);
                if (expression == null)
                {
                    throw new ArgumentException($"셀 '{cellName}'에 대한 메서드를 찾을 수 없습니다.");
                }

                double result = expression.Evaluate();

                // 결과 값이 무한대이거나 허용 범위를 초과하는 경우 OverflowException 발생
                if (double.IsInfinity(result) || Math.Abs(result) > MaxAllowedValue)
                {
                    throw new OverflowException($"계산 결과가 허용 범위를 초과했습니다: 셀 {cellName}, t={t}, 결과={result}");
                }

                sheet.SetValue(cellName, t, result);

            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(ex.Message);
            }
        }

        public void Clear()
        {
            CurrentSheetName = Name;
            Parameters = new Dictionary<string, object>();
            Sheets = new Dictionary<string, Sheet>();
        }
    }

    public class CompiledCell
    {
        public string Name { get; set; }
        public string Formula { get; set; }
        public string TransformedFormula { get; set; }
        public string Description { get; set; }
        public Model Model { get; set; }
        public IGenericExpression<double> Expression { get; set; }

        public bool IsCompiled { get; set; }
        public string CompileStatusMessage { get; set; }

        public CompiledCell(string name, string formula, string description, Model model)
        {
            Name = name;
            Formula = formula;
            Description = description;
            Model = model;
            Compile();         
        }

        private void Compile()
        {
            try
            {
                TransformedFormula = FormulaTransformationUtility.TransformText(Formula, Model.Name);
                Expression = Model.Engine.Context.CompileGeneric<double>(TransformedFormula);
                IsCompiled = true;
                CompileStatusMessage = "Successfully Compiled";
            }
            catch(Exception ex)
            {
                IsCompiled = false;
                CompileStatusMessage = ex.Message;
            }
        }
    }

}
