using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RoomInfo
{
    public static class KUtil
    {
        //壁の線分（Curve）を抽出するための関数。連続した2マス以上の線分を見つける
        public static IList<Curve> FindContinuousOnes(int[,] array, int a)
        {
            //結果として返す変数
            IList<Curve> curves = new List<Curve>();

            //行列のサイズを取得
            int rows = array.GetLength(0);
            int cols = array.GetLength(1);

            // 横方向の連続を確認
            for (int i = 0; i < rows; i++)
            {
                double startOffsetX = 0;
                double endOffsetX = 0;
                int start = -1;
                for (int j = 0; j < cols; j++)
                {
                    if (array[i, j] == a && start == -1)
                    {
                        // 連続の開始点を記録
                        start = j;
                        if (j >0)
                        {
                            if (a == 2 && array[i, j - 1] == 1)//******内壁を追跡中に左が外壁1だったら1つ減じる。（i-1>=0とする）
                            {
                                startOffsetX = -1;
                            }
                        }
                    }
                    else if (array[i, j] != a && start != -1)
                    {
                        // 連続が終了した場合
                        if (j - start > 1) // 2つ以上の連続のみ出力
                        {
                            //"横方向の連続はデータ配列上で: ({i}, {start}) から ({i}, {j - 1})"
                            //Revitモデル上では行はY軸、列はX軸にあたるので入れ替え。さらにY軸は大小の方向が逆なので
                            //最大値からiを引いている
                            //XYZ startPoint = new XYZ(i, start, 0);
                            //XYZ endPoint = new XYZ(i, j - 1, 0);

                            if (a == 2 && array[i, j] == 1)//******内壁を追跡中に右が外壁だったら1つ増やす。
                            {
                                endOffsetX = 1;
                            }

                            XYZ startPoint = new XYZ(start+ startOffsetX, -i ,  0);
                            XYZ endPoint = new XYZ((j - 1)+ endOffsetX, -i, 0);
                            Line line = Line.CreateBound(startPoint, endPoint);
                            curves.Add(line);
                            startOffsetX = 0;//オフセットのリセット
                            endOffsetX = 0;
                        }
                        start = -1; // 開始点をリセット
                    }
                }
                // 行末まで連続が続いた場合の処理
                if (start != -1 && cols - start > 1) // 行末までの連続が2つ以上の場合
                {
                    //"横方向の連続はデータ配列上で:  ({i}, {start}) から ({i}, {cols - 1})");
                    //Revitモデル上では行はY軸、列はX軸にあたるので入れ替え。さらにY軸は大小の方向が逆なので
                    //最大値からiを引いている
                    //XYZ startPoint = new XYZ(i, start, 0);
                    //XYZ endPoint = new XYZ(i, cols - 1, 0);
                    XYZ startPoint = new XYZ(start , -i,  0);
                    XYZ endPoint = new XYZ((cols - 1) , -i, 0);
                    Line line = Line.CreateBound(startPoint, endPoint);
                    curves.Add(line);
                }
            }

            // 縦方向の連続を確認
            for (int j = 0; j < cols; j++)
            {
                double startOffsetY = 0;
                double endOffsetY = 0;
                int start = -1;
                for (int i = 0; i < rows; i++)
                {
                    if (array[i, j] == a && start == -1)
                    {
                        // 連続の開始点を記録
                        start = i;
                        if (i > 0)
                        {
                            if (a == 2 && array[i-1, j] == 1)//******内壁を追跡中に上が外壁1だったら1つ減じる。（i-1>=0とする）
                            {
                                startOffsetY = -1;
                            }
                        }
                    }
                    else if (array[i, j] != a && start != -1)
                    {
                        // 連続が終了した場合
                        if (i - start > 1) // 2つ以上の連続のみ出力
                        {
                            //{start}, {j}) から ({i - 1}, {j})
                            //Revitモデル上では行はY軸、列はX軸にあたるので入れ替え。さらにY軸は大小の方向が逆なので
                            //最大値からiを引いている
                            //XYZ startPoint = new XYZ(start, j, 0);
                            //XYZ endPoint = new XYZ(i - 1, j, 0);

                            if (a == 2 && array[i, j] == 1)//******内壁を追跡中に下が外壁だったら1つ増やす。
                            {
                                endOffsetY = 1;
                            }

                            XYZ startPoint = new XYZ(j , -start- startOffsetY, 0);//方向が逆であることに注意
                            XYZ endPoint = new XYZ(j , -(i - 1) - endOffsetY, 0);//同上
                            Line line = Line.CreateBound(startPoint, endPoint);
                            curves.Add(line);
                        }
                        start = -1; // 開始点をリセット
                    }
                }
                // 列末まで連続が続いた場合の処理
                if (start != -1 && rows - start > 1) // 列末までの連続が2つ以上の場合
                {
                    //"横方向の連続はデータ配列上で:  ({start}, {j}) から ({rows - 1}, {j})");
                    //Revitモデル上では行はY軸、列はX軸にあたるので入れ替え。さらにY軸は大小の方向が逆なので
                    //最大値からiを引いている
                    //XYZ startPoint = new XYZ(start, j, 0);
                    //XYZ endPoint = new XYZ(rows - 1, j, 0);
                    XYZ startPoint = new XYZ(j , -start ,  0);
                    XYZ endPoint = new XYZ(j , -(rows - 1),  0);
                    Line line = Line.CreateBound(startPoint, endPoint);
                    curves.Add(line);
                }
            }
            return curves;
        }
        //Excelファイルを読むためのダイアログを表示する
        public static string OpenExcel()
        {
            string fileName = "";

            //OpenFileDialog
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "ファイル選択";
                openFileDialog.Filter = "ExcelFiles | *.xls;*.xlsx;*.xlsm";
                //初期表示フォルダはデスクトップ
                openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                //ファイル選択ダイアログを開く
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    fileName = openFileDialog.FileName;
                }
            }
            return fileName;
        }
        public static SaveFileDialog SaveExcel()
        {
            //Excelファイルに保存するためにファイル名をユーザーにきくダイアログを表示する
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "ExcelFiles | *.xls;*.xlsx;*.xlsm";
            saveFileDialog.Title = "EXCELファイル保存";
            saveFileDialog.ShowDialog();
            //ユーザーがファイル拡張子入れない場合にはxlsになってしまう。あとで開くときに警告が出るのであらかじめxlsxに変更しておく
            //filenameはユーザーが何も入れずにキャンセルボタンを押すと空になってしまう。
            String filename = saveFileDialog.FileName;
            //半角空白が含まれているとExcelアプリ起動時にファイルが見つからないエラーが出る場合あるので半角空白をアンダースコアに置き換える■
            filename = filename.Replace(" ", "_");
            if (filename.IndexOf("xlsx") < 0)
            {
                //SaveFileDialogオブジェクトのFileName属性も変更しておく
                saveFileDialog.FileName = filename.Split('.')[0] + ".xlsx";
            }
            return saveFileDialog;
        }
        //ミリメーターを内部単位に変換する
        public static double CVmmToInt(double x)
        {
            return UnitUtils.ConvertToInternalUnits(x, UnitTypeId.Millimeters);
        }
        //内部単位をミリメーターに変換する
        public static double CVintTomm(double x)
        {
            return UnitUtils.ConvertFromInternalUnits(x, UnitTypeId.Millimeters);
        }
        //窓やドアを取り付ける
        public static void CreateWindowAndDoor(UIDocument uidoc, Document doc, string fsFamilyName,
            string fsName, string levelName, double xCoord, double yCoord, double offset)
        {
            // LINQ 指定のファミリシンボルを探して取得する
            FamilySymbol familySymbol = (from fs in new FilteredElementCollector(doc).
                 OfClass(typeof(FamilySymbol)).
                 Cast<FamilySymbol>()
                                         where (fs.Family.Name == fsFamilyName && fs.Name == fsName)
                                         select fs).First();

            // LINQ 指定のレベルを取得する
            Level level = (from lvl in new FilteredElementCollector(doc).
                           OfClass(typeof(Level)).
                           Cast<Level>()
                           where (lvl.Name == levelName)
                           select lvl).First();

            // 座標mmを内部単位に変換する
            double x = xCoord;
            double y = yCoord;

            XYZ xyz = new XYZ(x, y, level.Elevation+ offset);

            //部材を挿入する壁で、最も近いホスト壁を取得する
            //壁を全部取得する
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(Wall));
            //当該レベルにある壁だけ抽出する。リストにする
            List<Wall> walls = collector.Cast<Wall>().Where(wl => wl.LevelId == level.Id).ToList();

            Wall wall = null;
            //距離の初期値を最大値にしておく
            double distance = double.MaxValue;
            //壁を全部調べて、最も近いホスト壁を抽出する
            foreach (Wall w in walls)
            {
                //壁のカーブからの最短直線距離を測る
                double proximity = (w.Location as LocationCurve).Curve.Distance(xyz);
                //これまでの値よりも小さい場合には、より接近した壁だと思われるので、その壁をホスト候補とする
                if (proximity < distance)
                {
                    distance = proximity;
                    wall = w;
                }
            }

            // Create window.
            //using (Transaction t = new Transaction(doc, "Create window and door"))
            //{
                //t.Start();

                if (!familySymbol.IsActive)
                {
                    //ファミリシンボルがアクティブではないので、アクティブにする
                    familySymbol.Activate();
                    doc.Regenerate();
                }

                // 部材の配置
                // ホストであるwallを指定しない場合には部材はホストなし
                FamilyInstance window = doc.Create.NewFamilyInstance(xyz, familySymbol, wall, level,StructuralType.NonStructural);
                //t.Commit();
            //}
            //string prompt = "部材は配置されました";
            //TaskDialog.Show("Revit", prompt);
        }
        //目的のレベル（階）を探して返す
        public static Level GetLevel(Document doc, string levelName)
        {
            Level result = null;
            //エレメントコレクターの作成
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            //レベルを全て検出する
            ICollection<Element> collection = collector.OfClass(typeof(Level)).ToElements();
            //目的のレベルを探す
            foreach (Element element in collection)
            {
                Level level = element as Level;
                if (null != level)
                {
                    if (level.Name == levelName)
                    {
                        result = level;
                    }
                }
            }
            return result;
        }
        //床作成のための外壁座標を取得する（数字が入っているセルの座標を取得する）
        public static SortedDictionary<int, XYZ> FindFloorLine(string[,] array)
        {
            //自動ソートされるディクショナリ
            var dict = new SortedDictionary<int, XYZ>();
            //行列のサイズを取得
            int rows = array.GetLength(0);
            int cols = array.GetLength(1);

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    int numeric;
                    if (int.TryParse(array[i, j], out numeric))
                    {
                        //dict.Add(numeric, new XYZ(i, j, 0));
                        dict.Add(numeric, new XYZ(j , -i, 0));
                    }
                }
            }
            return dict;
        }
        //床のためのカーブを作る
        public static IList<XYZ> FindSymbolPosition(string[,] array, string symbol)
        {
            //自動ソートされるディクショナリ
            var result = new List<XYZ>();
            //行列のサイズを取得
            int rows = array.GetLength(0);
            int cols = array.GetLength(1);

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    if(array[i, j] != null)
                    {
                        //シンボルと一致したらリストに入れる
                        if (array[i, j].Equals(symbol))
                        {
                            result.Add(new XYZ(j, -i, 0));
                        }
                    }
                }
            }
            return result;
        }
        public static void CreateColumn(Document doc, String levelName1,string levelName2,string columnFamilyName, string columnSymbolName,double x,double y)
        {
            string message = null;

                try
                {
                // 柱のファミリとタイプを取得
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                    FamilySymbol columnSymbol = collector
                        .OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
                        .FirstOrDefault(q => q.Name == columnSymbolName && q.Family.Name == columnFamilyName);

                    if (columnSymbol == null)
                    {
                        message = "指定された柱のファミリまたはタイプが見つかりません。";
                    }

                    // レベルを取得
                    Level level1 = GetLevel(doc, levelName1);
                    Level level2 = GetLevel(doc, levelName2);

                    if (level1 == null || level2 == null)
                    {
                        message = "指定されたレベルが見つかりません。";
                    }

                    // トランザクションを開始
                    //using (Transaction trans = new Transaction(doc, "Place Column"))
                    //{
                        //trans.Start();

                        // ファミリシンボルをアクティブにする
                        if (!columnSymbol.IsActive)
                        {
                            columnSymbol.Activate();
                        }

                        // 柱を作成
                        XYZ location = new XYZ(x, y, 0);
                        FamilyInstance column = doc.Create.NewFamilyInstance(location, columnSymbol, level1, Autodesk.Revit.DB.Structure.StructuralType.Column);

                        // 上部の拘束を設定
                        
                        Parameter topLevelParam = column.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM);
                        if (topLevelParam != null && !topLevelParam.IsReadOnly)
                        {
                            topLevelParam.Set(level2.Id);
                        }
                        
                        //trans.Commit();
                    //}

                }
                catch (Exception ex)
                {
                    message = ex.Message;
                }
            if (message != null)
            {
                TaskDialog.Show("Error", message);
            }
        }
        public static void CreateGrids(Document doc)
        {
            double y1 = -100;
            double y2 = 30;
            double[] x = {2,34,74,114,147 };
            string[] symbX = {"1","2","3","4","5" };

            double x1 = -30;
            double x2 = 180;
            double[] y = { -2, -33, -66 };
            string[] symbY = { "A", "B", "C"};

            for (int i = 0; i < x.Length; i++) 
            {
                XYZ start = new XYZ(x[i], y1, 0);
                XYZ end = new XYZ(x[i], y2, 0);
                Line line = Line.CreateBound(start, end);
                Grid grid = Grid.Create(doc, line);
                grid.Name = symbX[i];
            }

            for (int i = 0; i < y.Length; i++)
            {
                XYZ start = new XYZ(x1, y[i], 0);
                XYZ end = new XYZ(x2, y[i], 0);
                Line line = Line.CreateBound(start, end);
                Grid grid = Grid.Create(doc, line);
                grid.Name = symbY[i];
            }
        }
        public static bool IsValidView(Autodesk.Revit.DB.View view)
        {
            if (view?.Category == null) return false;
            return view.Category.Name.Equals("ビュー") || view.Category.Name.Equals("Views");
        }
    }
}
