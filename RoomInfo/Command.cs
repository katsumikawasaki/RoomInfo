#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.IFC;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Drawing;
using Application = Autodesk.Revit.ApplicationServices.Application;

#endregion
//必要な参照
/*
*EPPlus ver4.5.3.1 これはLGPLライセンスです。著作権はEPPlus Software社です。
*Microsoft
*PresentationCore
*PresentationFramework
*RevitAPI
*RevitAPIIFC
*RevitAPIUI
*System
*System.configuration
*System.Data
*System.Drawlng
*Syatem.Security
*System.Windows
*System.Windows.Foms
*System.Xaml
*WindowsBase
*WindowsFormsIntegration
*
*/
//ファイル構成
/*
 * Command.cs
 * package.config
 * ProcessStatusUI.xaml
 *   ProcessStatusUI.xaml.cs
 * RoomInfo.aain
 */
namespace RoomInfo
{
    [Transaction(TransactionMode.Manual)]
	public class Command : IExternalCommand
	{
		private Document doc;
        private FilteredElementCollector roomCollection;
		private List<RoomInformation> listRooms;


		public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
		{
			//Revitバージョンをチェックする。例＝2022


            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;

            var version = app.VersionNumber;
			if (!version.Equals("2022"))
			{
				TaskDialog.Show("エラー", "このアドインはRevit2022対応です。しかしこのRevitは" + version + "です。");
				return Result.Failed;
			}
			try
			{
				//現在の画面がビューであるときにこのアドインが実行できる。それをチェック
				Autodesk.Revit.DB.Category actView = doc.ActiveView.Category;
				if (actView != null)
				{
					if (actView.Name.Equals("ビュー") || actView.Name.Equals("Views"))//日英に対応
					{
						//TaskDialog.Show("OK","現在ビュー画面です");//
					}
					else
					{
						TaskDialog.Show("エラー", "ビュー画面(3D、平面図、立面図など)表示してからアドインを実行してください");
						return Result.Failed;
					}
				}
				else
				{
					TaskDialog.Show("エラー", "3Dビューを表示してからアドインを実行してください");
					return Result.Failed;
				}
				//アクティブなビューの取得
				var activeViewGraphic = commandData.Application.ActiveUIDocument.ActiveGraphicalView;
                //カレントフェーズの取得とフェーズフィルター(現在のフェーズに合ったフィルター)
                var phaseProvider = new ParameterValueProvider(new ElementId(BuiltInParameter.ROOM_PHASE));
                var currentPhase = activeViewGraphic.get_Parameter(BuiltInParameter.VIEW_PHASE).AsElementId();
				var phaseRule = new FilterElementIdRule(phaseProvider, new FilterNumericEquals(), currentPhase);
				ElementParameterFilter phaseFilter = new ElementParameterFilter(phaseRule);
				//面積がゼロの部屋を除外するフィルター。モデルには以前に作成して削除された部屋も残っている。面積がゼロで。それを除外する
				ParameterValueProvider areaProvider = new ParameterValueProvider(new ElementId(BuiltInParameter.ROOM_AREA));
                ElementParameterFilter areaFilter = new ElementParameterFilter(new FilterDoubleRule(areaProvider, new FilterNumericGreater(), 0, 0.0001));
                //上記2つのフィルターをANDでつなぐ
                LogicalAndFilter andFilter = new LogicalAndFilter(phaseFilter, areaFilter);
                //上記のフィルターをかけて部屋一覧を取得する。
                roomCollection = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WherePasses(andFilter);
				
				//RoomInFomation型のリストを新規作成。室の情報を保存するためのリスト
				listRooms = new List<RoomInformation>();

				//これからプログレスバーを表示して時間がかかる処理を行う
				//
				//全部の室数の取得
				int full = roomCollection.Count();
				//プログレスバーの表示と時間がかかる処理の実行
				ProgressStatusUI progressBarUI = new ProgressStatusUI();
                //プログレスバーの大きさ指定
                progressBarUI.Width = 450;
				progressBarUI.Height = 180;
				//プログレスバーが表示されたら実行する関数を指定する。関数ProgressBarUI_ContentRenderedはこの下のほうに有る。
				progressBarUI.ContentRendered += ProgressBarUI_ContentRendered;
				//プログレスバー表示を中央に。
				var desktop = System.Windows.Forms.Screen.PrimaryScreen.Bounds;//デスクトップサイズ
				progressBarUI.Top = Convert.ToInt32(((double)desktop.Height - (double)progressBarUI.Height) / 2.0);
				progressBarUI.Left = Convert.ToInt32(((double)desktop.Width - (double)progressBarUI.Width) / 2.0);
				progressBarUI.ShowDialog();
				//ユーザーがキャンセルをしたかどうかをチェック
				if (progressBarUI.checkCancel())
				{
					//キャンセルした場合はaddinもキャンセル終了
					return Result.Cancelled;
				}
                //プログレスバーが進みながら関数ProgressBarUI_ContentRenderedが実行される
                //ProgressBarUI_ContentRenderedの実行が終了したらこの次の命令から実行される

                bool noVolumeError = true;//全部の部屋の面積がゼロだったらtrueになる
				foreach(RoomInformation room  in listRooms)
				{
					if(room.Volume != 0.0)
					{
						noVolumeError = false;//面積が1つでもゼロでないならfalseにする
					}
				}
				if(noVolumeError)
				{
					//全室の面積がゼロの場合はエラー出す
					TaskDialog.Show("OK", "部屋が一つも設定されていないか、Revitの設定で部屋の容積計算を行うことが設定されていないようです。部屋があることと、設定で「面積と容積計算をする」選択になっていることを確認して再度このアドインを実行してください");
				}

                //換気計算書に室情報を転記するか
                //ユーザーに聞く
                TaskDialog taskDialog = new TaskDialog("換気計算書への転記");
                taskDialog.MainContent = "Revitから取得した室情報を換気計算書に転記しますか？Yesは転記します、Noは行わないです。換気計算書Excelはこの後ファイル選択していただきます。";
                taskDialog.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No;
                taskDialog.DefaultButton = TaskDialogResult.No;
                if (taskDialog.Show() == TaskDialogResult.Yes)
                {
                    string kankiExcel = KUtil.OpenExcel();
                    if (kankiExcel != "")
                    {
                        //Excelファイルストリームを取得
                        FileInfo finfo = new FileInfo(kankiExcel);
                        using (ExcelPackage package2 = new ExcelPackage(finfo))
                        {
                            //DataTableを作成する
                            var kankitable = new DataTable("kankiData");
                            //列名、型を設定する
                            kankitable.Columns.Add("階", typeof(string));//0
                            kankitable.Columns.Add("室名", typeof(string));//1
                            kankitable.Columns.Add("室番号", typeof(string));//2
                            kankitable.Columns.Add("床面積(m2)", typeof(double));//3
                            kankitable.Columns.Add("天井高(m)", typeof(double));//4

                            for (int i = 0; i < listRooms.Count; i++)
                            {
                                var row = kankitable.NewRow();
                                row[0] = listRooms[i].Level.Name;//階
                                row[1] = listRooms[i].Name;//室名
                                row[2] = listRooms[i].Number;//室番号
                                row[3] = Math.Round(UnitUtils.ConvertFromInternalUnits(listRooms[i].Area, UnitTypeId.SquareMeters), 3);
                                row[4] = Math.Round(UnitUtils.ConvertFromInternalUnits(listRooms[i].Height, UnitTypeId.Millimeters)/1000, 3);
                                kankitable.Rows.Add(row);
                            }
                            ExcelWorksheet kankiworksheet = package2.Workbook.Worksheets[1];//ワークシートの番号はOfficeによって違うので注意
                            //Excelの一番左上の位置C6からDataTableを一気に流し込む
                            kankiworksheet.Cells["A7"].LoadFromDataTable(kankitable, false);//ヘッダー無しで記入

                            //Excelファイルを保存する
                            package2.Save();
                        }
                        TaskDialog.Show("正常終了", "換気計算書に書き込みました");
                    }
                }
			}
			catch (Exception ex)//例外発生した場合にはキャッチしてエラーメッセージを出す。
			{
				TaskDialog.Show("エラー", "処理は失敗しました。原因=" + ex.Message);
			}
			return Result.Succeeded;
		}

		//1室の内法面積を計算する
		private double CalculateRoomInnerMenseki(Room room)
		{
			//この関数から返す値。内法面積
			double result = 0.0;
			//室面積がゼロのものは、一旦削除された室なので、無視する。リターンする
			if (room.Area < 0.0001) return result;

			//次のバウンダリを取得するためのオプションを作る
			SpatialElementBoundaryOptions opts = new SpatialElementBoundaryOptions();
			//バウンダリの要素、つまりほとんど壁だと想定しているもの。（壁ではないものも可能性としてある）
			IList<IList<BoundarySegment>> loops = room.GetBoundarySegments(opts);
			
			//バウンダリ（面積を構成する複数の境界線）を分析する
			//一室分の壁の座標を保存するための配列, 一応100個用意する
			double[,] points = new double[100, 100];

			foreach (IList<BoundarySegment> loop in loops)
			{
				//座標をいくつ使ったかのインデックス。初期化
				int pointIndex = 0;
				//endPointを一時的に記憶するための変数。初期化
				XYZ lastEndPoint = null;

				//一つのループについて詳しく分析する
				foreach (BoundarySegment boundarySegment in loop)
				{
					//壁のエレメントID
					ElementId idWall = boundarySegment.ElementId;
					//壁（roomに面する仕上げ面でZ軸はフロアレベルになっている）の線分の開始点と終了点を調べる（これは壁の内法に沿った線分である。
					//本プログラムは直線の壁しか計算できない）もし曲線があると直線で計算してしまうので誤差が出る
					//バウンダリセグメントから接している部分の壁のCurveを取得する
					Curve curve = boundarySegment.GetCurve();
					Arc arc = null;
					try {
						arc = (Arc)curve;
					}catch(Exception e)
                    {
						string dummy = e.Message;
                    }
					XYZ startPoint = curve.GetEndPoint(0);
					XYZ endPoint = curve.GetEndPoint(1);

					//スタートポイントの座標を面積計算のために格納しておく。反時計回りに格納する必要がある。
					points[0, pointIndex] = startPoint.X;
					points[1, pointIndex] = startPoint.Y;
					//endPointは一時的に覚えて置く。最後に使う
					lastEndPoint = endPoint;
					//pointIndexを増やしておく
					pointIndex += 1;

				}//１つのループ（いくつかの壁でかこまれた１つのエリア）の処理の終わり
				 //１つのループごとに、その面積を計算して加減算する。

				 //スタートポイントの座標を面積計算のために格納しておく
				 //データのポイントは反時計回りに格納する必要がある。最後はスタートポイントになっていること。！！
				points[0, pointIndex] = lastEndPoint.X;
				points[1, pointIndex] = lastEndPoint.Y;
				//pointIndexを増やしておく
				pointIndex += 1;

				//roomの内法面積を計算して加算する。例えば室の内部に独立柱があるときは
				//マイナスの面積がtakakukeiMenseki関数から返されるので差し引かれることになる
				result += TakakukeiMenseki(points, pointIndex);

			}//全部のループの処理の終わり

			return result;
		}

		private double TakakukeiMenseki(double[,] points, int n)
		{
			//この面積計算方法は2次元上の多角形の面積を計算する「くつひも公式」を使用している。
			//2つのベクトルを2辺とするひし形の面積はベクトルの外積の大きさになるということに基づいている。
			//pointsは、例えば四角形の頂点A,B,C,Dが半時計周り並ばなければいけない。
			//最後にAの座標に戻るように付け加えて、次のような配列になるようにする
			// [Ax, Bx, Cx, Dx, Ax]
			// [Ay, By, Cy, Dy, Ay]
			//ポイント数nは横軸の要素数で、上記の例だと5になる

			//ポイント数nが3個の場合には円柱など円形を意味していることが考えられる
			//1番目と2番目のポイントは円の左端と右端を意味しているので面積を計算して、マイナスで返す
			if (n == 3)
			{
				return Math.PI * ((points[0,1]-points[0,0])* (points[0,1] - points[0,0])) * (-0.25);
			}

			//多角形の面積
			double result = 0.0;
			for (int i = 0; i < n - 1; i++)
			{
				result += points[0, i] * points[1, i + 1] - points[1, i] * points[0, i + 1];
			}
			result *= 0.5;
			//面積を返す
			return result;
		}

		//プログレスバーが表示された直後に実行される関数。この中で時間がかかる処理を行う
		private void ProgressBarUI_ContentRendered(object sender, EventArgs e)
		{
			ProgressStatusUI progressBarUI = sender as ProgressStatusUI;

			if (progressBarUI == null)
				throw new Exception("ステータスバー作成のときにエラーが発生しました");

			int numberOfRooms = roomCollection.Count();
			int roomCount = 1;

			//各室に関してループして分析する
			foreach (Room room in roomCollection.Cast<Room>())
			{
                //1室に対して1つのRoomInformationオブジェクトを新規作成
				//これに各プロパティを入れていく。
                RoomInformation roomInfor = new RoomInformation(room);

				//部屋の面積を内法で計算してセットする。容積から高さを算出するため
				roomInfor.InnerArea = CalculateRoomInnerMenseki(room);

				//容積を内法面積で割ることで高さHeight（平均値）を計算
				//容積は常に正しい値を返すので利用する
				if (roomInfor.InnerArea != 0.0)
				{
					roomInfor.Height = roomInfor.Volume / roomInfor.InnerArea;
				}
				else
				{
					roomInfor.Height = 0.0;
				}

                //RoomInformation型のリストに以上の1室分のデータを追加する
                listRooms.Add(roomInfor);


				/////プログレスバーのステータス更新
				int progressPercent = Convert.ToInt32((double)roomCount / (double)numberOfRooms * 100.0);
				progressBarUI.UpdateStatus(string.Format("処理した室数 {0}", roomCount.ToString()) + "/" + numberOfRooms.ToString(), progressPercent);
				if (progressBarUI.ProcessCancelled)
					break;

                roomCount++;
			}
			//室リストを階順に並べ替える。これは地盤面からの高さを意味するElevationの値によって小さい順に並べ替える。
			listRooms = listRooms.OrderBy(x => x.Level.Elevation).ThenBy(x => x.Number).ToList();

			//階番号を意味するLevel Idカラム（LevelId）を単純な数値に置き換える。
			//階名称を意味するLevelカラム（Level.Name）と階番号の関係を調べる
			//辞書型変数の用意
			Dictionary<string, double> levelTable = new Dictionary<string, double>();
			for (int i = 0; i < listRooms.Count; i++) {
				string levelName = listRooms[i].Level.Name;
				if (!levelTable.ContainsKey(levelName))
				{
					levelTable.Add(levelName, listRooms[i].Level.Elevation);//その階のGLからの高さを値として登録
				}
			}
			//階が低いものから高くなる順に並べる。Valueには上記の処理によってElevationの数値が入っているので小さい順に並べ替える
			levelTable.OrderByDescending(x => x.Value);

			//値を単純な整数化する。例えば地下1階は-1、1階は0、2階は1とする
			//１階の整数は記憶しておく
			//この処理のためにlevelTableをList型に変換する
			List<KeyValuePair<string, double>> levelList = levelTable.ToList();
			//地階の数のカウンター
			int underGroundFloorCount = 0;
			for (int i=0;i<levelList.Count-1;i++)
			{
				if(levelList[i].Value < 0.0)
				{
					//建築基準法の定義に準じてi番目のフロアの階高（法律では天井高さだがそれを計算するのが複雑になるため近似的に計算している）の3分の1以上の高さが地下に埋まっている階を地階とする
					if(Math.Abs(levelList[i+1].Value - levelList[i].Value)/3 <= Math.Abs(levelList[i].Value))
					{
						//この場合にはフロアiは地階であるから地階数カウントを増やす
						underGroundFloorCount++;
					}
				}
			}
			//最初のフロアの階を決める（=地階の数）
			int renumberStart = -1 * underGroundFloorCount;
			//整数で振りなおすためのディクショナリ変数
			Dictionary<string, int> newLevelTable = new Dictionary<string, int>();
			foreach (KeyValuePair<string, double> data in levelTable)
			{
				//1階がゼロになるように番号を振りなおしたものを新規で作る
				newLevelTable.Add(data.Key, renumberStart);
				renumberStart++;//階があがるごとに１増やす。1階はゼロになるようになっている
			}

			//データテーブルのlistRoomsのLevelSequenceNumber番号を設定する
			for (int i = 0; i < listRooms.Count; i++)
            {
				listRooms[i].LevelSequenceNumber = newLevelTable[listRooms[i].Level.Name];
			}

			//プログレスバーを閉じる
			progressBarUI.Close();
		}
	}

	public class RoomInformation
	{
        //室名
        public string Name { set; get; }
        //室番号
        public string Number { set; get; }
        //容積
        public double Volume { set; get; }
        //面積
        public double Area { set; get; }
        //室の内法面積
        public double InnerArea { set; get; }
        //要素ID
        public ElementId Id { set; get; }
        //階（レベル）
        public Level Level { set; get; }
        //上記のId
        public ElementId LevelId { set; get; }
        //高さ
        public double Height { set; get; }
        //階の連番情報
        public int LevelSequenceNumber { set; get; }

        public RoomInformation(Room room)
		{
			Name = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString();//room.Nameでもよいが、その場合には室Numberも加わってしまう
			Name = Name.Replace("\n", "");//念のため改行を削除する
			Number = room.Number;
			Volume = room.Volume;
			Area = room.Area;
			Id = room.Id;
			Level = room.Level;
			LevelId = room.LevelId;
		}
	}
}