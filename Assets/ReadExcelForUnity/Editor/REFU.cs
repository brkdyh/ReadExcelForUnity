using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using UnityEditor;
using UnityEngine;
using System;

namespace REFU
{
    public class REFU : EditorWindow
    {

        static REFU instance;
        [MenuItem("REFU/Open Window")]
        public static void OpenWindow()
        {
            if (instance == null)
            {
                instance = CreateWindow<REFU>();
                instance.minSize = new Vector2(480, 200);
                instance.title = "REFU Window";
                instance.exportPath = Application.dataPath + "/Resources/REFU";
            }

            instance.Show();
            instance.Focus();
        }

        string excelPath;
        string exportPath;
        private void OnGUI()
        {
            GUILayout.Space(10);
            GUILayout.BeginHorizontal("box");
            GUILayout.Label("当前选择文件:");
            GUILayout.Label(Path.GetFileName(excelPath));

            GUILayout.FlexibleSpace();

            if (GUILayout.Button("选择表格文件"))
            {
                excelPath = EditorUtility.OpenFilePanel("选择Excel表格", Application.dataPath.Replace("/Assets", ""), "xlsx,xls");
            }

            GUILayout.EndHorizontal();


            if (EditorApplication.isCompiling)
            {
                EditorGUILayout.HelpBox("请等待编译完成", MessageType.Warning);
                return;
            }

            GUILayout.Space(10);


            if (excelPath == null)
                return;

            GUILayout.BeginVertical("box");
            GUILayout.BeginHorizontal();
            GUILayout.Label("当前输出路径:");
            GUILayout.Label(exportPath);
            GUILayout.EndHorizontal();

            GUILayout.Space(10);
            if (GUILayout.Button("选择输出路径"))
            {
                exportPath = EditorUtility.OpenFolderPanel("选择路径", Application.dataPath.Replace("/Assets", ""), "");
            }
            GUILayout.Space(10);

            GUILayout.EndVertical();

            GUILayout.FlexibleSpace();

            GUILayout.BeginHorizontal("box");
            GUILayout.Space(10);
            if (GUILayout.Button("生成数据类型"))
            {
                ClearLoad();
                using (var excel = LoadExcel(excelPath))
                {
                    foreach (var sheet in _excelWorksheet)
                    {
                        var tfi = getFieldInfo(sheet.Value);
                        CodeGenerator.CreateType(sheet.Key, tfi);
                    }
                }
            }

            if (GUILayout.Button("读取表格数据"))
            {
                if (!Directory.Exists(exportPath))
                    Directory.CreateDirectory(exportPath);

                ClearLoad();
                using (var excel = LoadExcel(excelPath))
                {
                    foreach (var sheet in _excelWorksheet)
                    {
                        var tfi = getFieldInfo(sheet.Value);
                        LoadSheet(sheet.Value, tfi);
                    }
                }

                AssetDatabase.SaveAssets();
                AssetDatabase.Refresh();
            }

            GUILayout.Space(10);
            GUILayout.EndHorizontal();

            GUILayout.Space(20);

            Repaint();
        }

        Dictionary<string, ExcelPackage> _excelPackages = new Dictionary<string, ExcelPackage>();
        Dictionary<string, ExcelWorksheet> _excelWorksheet = new Dictionary<string, ExcelWorksheet>();

        private void ClearLoad()
        {
            _excelPackages.Clear();
            _excelWorksheet.Clear();
        }

        private ExcelPackage LoadExcel(string excel)
        {
            if (!_excelPackages.ContainsKey(excel))
            {
                var excelInfo = new FileInfo(excel);
                if (!excelInfo.Exists)
                {
                    Debug.Log("该路径下找不到Excel文件： " + excel);
                    return null;
                }


                var _excel = new ExcelPackage(excelInfo);
                _excelPackages.Add(excel, _excel);

                var il = _excel.Workbook.Worksheets.GetEnumerator();
                while (il.MoveNext())
                {
                    var sheet = il.Current as ExcelWorksheet;
                    _excelWorksheet.Add(sheet.Name, sheet);
                }
            }

            return _excelPackages[excel];
        }

        private ExcelWorksheet GetWorksheet(string sheet)
        {
            if (_excelWorksheet.ContainsKey(sheet))
                return _excelWorksheet[sheet];

            return null;
        }

        TypeFieldInfo[] getFieldInfo(ExcelWorksheet sheet)
        {
            if (sheet.Dimension == null)
                return null;


            var col_count = sheet.Dimension.Columns;

            TypeFieldInfo[] fis = new TypeFieldInfo[col_count];
            for (int col = 1; col <= col_count; col++)
            {
                var name = sheet.GetValue<string>(1, col);
                //Debug.Log(name);
                var type = sheet.GetValue<string>(2, col);
                //Debug.Log(type);

                fis[col - 1] = new TypeFieldInfo();
                fis[col - 1].FieldName = name;
                fis[col - 1].FieldType = TypeMapper.TYPE_MAPPER.ContainsKey(type) ?
                     TypeMapper.TYPE_MAPPER[type] : System.Type.GetType(type);
                //Debug.Log("type = " + fis[col - 1].FieldType);
            }

            return fis;
        }

        //加载表格，反射赋值
        void LoadSheet(ExcelWorksheet sheet, TypeFieldInfo[] fieldInfos)
        {
            if (sheet == null)
            {
                Debug.LogError("Null Sheet!");
                return;
            }

            //Debug.Log(typeof(Person).Name);
            //Debug.Log(System.Type.GetType("Person"));
            //var type = System.Type.GetType(sheet.Name);
            //if (type == null)
            //{
            //    Debug.LogError("Can't Find Mapping Type by Sheet : " + sheet.Name);
            //    return;
            //}
            var data = ScriptableObject.CreateInstance(sheet.Name);

            if (data == null)
            {
                Debug.LogError("Can't Find Mapping Type by Sheet : " + sheet.Name);
                return;
            }

            var type = data.GetType();
            if (sheet.Dimension == null)
            {
                Debug.LogWarning("Sheet Dimension is Null : " + sheet.Name);
                return;
            }

            for (int col = 1; col <= sheet.Dimension.Columns; col++)
            {
                var field_name = fieldInfos[col - 1].FieldName;
                var field_type = fieldInfos[col - 1].FieldType;

                if (field_type == null)
                    continue;

                var get_field = type.GetField(field_name);

                if(get_field==null)
                {
                    Debug.LogError("Type Don't Contain Field,Try Re-Generate Code!");
                    return;
                }
                //if (get_field.FieldType != field_type)
                //{
                //    Debug.LogError("Field Type Can't Map! " + sheet.Name + " : " + field_name + " " + field_type + " <=> " + get_field.FieldType);
                //}
                var row_count = sheet.Dimension.Rows;
                Array array = Array.CreateInstance(field_type, row_count - 2);
                for (int row = 3; row <= row_count; row++)
                {
                    object value = sheet.GetValue(row, col);
                    try
                    {
                        value = Convert.ChangeType(value, field_type);
                    }
                    catch (Exception ex)
                    {
                        Debug.LogError("Convert Change Type Faild: " + ex.Message + "\n" + ex.StackTrace);
                        break;
                    }
                    array.SetValue(value, row - 3);
                }

                get_field.SetValue(data, array);
            }


            var path = exportPath.Replace(Application.dataPath, "Assets/");
            AssetDatabase.CreateAsset(data, path + "/" + sheet.Name + ".asset");
        }
    }
}