using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.DependencyInjection;
using ProjectApp.Utils;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Line = Autodesk.Revit.DB.Line;

namespace ProjectApp.ModelFromCad
{
    /// <summary>
    /// Interaction logic for ModelFromCadView.xaml
    /// </summary>
    public partial class ModelFromCadView : Window
    {
        private readonly Document document;
        public XyzData CadBeamOrigin;
        private ObservableCollection<BeamInfoCollection> _beamInfoCollections = new();
        private readonly List<ColumnInfoCollection> columnInfoCollections = new();
        private readonly List<CadBeams> _cadBeams = new();
        private XYZ _origin;
        public List<CadRectangle> cadRectangles = new();
        public List<List<XyzData>> ListPoint = new();
        public ObservableCollection<FloorInfoCollection> FloorInfoCollections { get; set; } = new();
        public ObservableCollection<BeamInfoCollection> BeamInfoCollections
        {
            get => _beamInfoCollections;
            set { _beamInfoCollections = value; }
        }
        public List<FloorInfoCollection> floorInfoCollections = new();

        public ObservableCollection<ColumnInfoCollection> ColumnInfoCollections { get; set; } = new();

        public ModelFromCadView(Document document)
        {
            InitializeComponent();
            this.document = document;
            LoadColumnFamilySymbols();
        }

        private void ButtonBase_OnClick(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void LoadColumnFamilySymbols()
        {
            //Lấy các column types trong dự án
            var columnFamilySymbols = new FilteredElementCollector(AC.Document)
                .OfCategory(BuiltInCategory.OST_StructuralColumns).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
                .Select(x => x.Family).Where(x => x.StructuralMaterialType != StructuralMaterialType.Steel)
                .DistinctBy(x => x.Id).OrderBy(x => x.Name).ToList();

            cboColumnFamilyType.ItemsSource = columnFamilySymbols;

            // Chọn item đầu tiên nếu có
            if (columnFamilySymbols.Any())
            {
                cboColumnFamilyType.SelectedIndex = 0;
            }

            // Lấy tất cả các element thuộc category Levels và class Level
            var levels = new FilteredElementCollector(document)
                .OfCategory(BuiltInCategory.OST_Levels)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .OrderBy(x => x.Elevation)
                .ToList();

            cbBaseLevel.ItemsSource = levels;
            cbBaseLevel.SelectedItem = levels.FirstOrDefault();
            cbTopLevel.ItemsSource = levels;
            cbTopLevel.SelectedItem = levels.Skip(1).FirstOrDefault();

            GetDataDefaultBeam();
            GetDataDefaultFloor();
        }

        private void ButtonBase_OnClick2(object sender, RoutedEventArgs e)
        {
            switch (TabControl.SelectedIndex)
            {
                case 1:
                    ModelBeams();
                    break;
                case 0:
                    ModelColumn();
                    break;
                default:
                    ModelFloor();
                    break;
            }
        }

        //Nút nhấn lấy dữ liệu từ file autocad
        private void ButtonBase_OnClick3(object sender, RoutedEventArgs e)
        {
            switch (TabControl.SelectedIndex)
            {
                case 1:
                    SelectBeam();
                    break;
                case 0:
                    SelectFormCadColumn();
                    break;
                default:
                    SelectFloorFromCad();
                    break;
            }
        }

        #region Beams

        // Lấy dữ liệu dầm từ autocad
        public void SelectBeam()
        {
            //beamInfoCollections.Clear();
            Hide();

            dynamic a = ComRunningObject.GetActiveObjectByProgId("AutoCaD.Application");
            a.Visible = true; // bật hiển thị
            a.WindowState = 3; // thường là Maximized (tùy enum COM)
            a.ActiveDocument.Activate(); // kích hoạt document
            dynamic doc
                = a.Documents.Application.ActiveDocument;

            dynamic b = a.Documents.Application.ActiveDocument;

            string[] arrPoint = null;

            try
            {
                var pointCad = doc.Utility.GetPoint(Type.Missing, "Select point: ");
                arrPoint = ((IEnumerable)pointCad).Cast<object>()
                    .Select(x => x.ToString())
                    .ToArray();
            }
            catch (Exception e)
            {
            }

            if (arrPoint != null)
            {
                double[] arrPoint1 = new double[3];

                int i = 0;

                foreach (var item in arrPoint)
                {
                    arrPoint1[i] = Convert.ToDouble(item);
                    i++;
                }

                CadBeamOrigin = new XyzData(arrPoint1[0], arrPoint1[1], arrPoint1[2]);
                var newset = doc.SelectionSets.Add(Guid.NewGuid().ToString());
                newset.SelectOnScreen();
                if (newset.Count <= 0)
                {
                }

                List<dynamic> listText = new List<dynamic>();

                List<dynamic> listLine = new List<dynamic>();

                //cot tron
                foreach (dynamic s in newset)
                {
                    if (s.EntityName == "AcDbLine")
                    {
                        listLine.Add(s);
                    }

                    if (s.EntityName == "AcDbText")
                    {
                        listText.Add(s);
                    }
                }

                List<TextData> listpoint = new List<TextData>();
                if (listText.Count > 0)
                {
                    foreach (var text in listText)
                    {
                        string[] arrtextpoint = ((IEnumerable)text.InsertionPoint).Cast<object>()
                            .Select(x => x.ToString())
                            .ToArray();

                        double[] arrtextpoint1 = new double[3];
                        int k = 0;
                        foreach (var item in arrtextpoint)
                        {
                            arrtextpoint1[k] = Convert.ToDouble(item);
                            k++;
                        }

                        listpoint.Add(new TextData()
                        {
                            point = new XYZ(arrtextpoint1[0], arrtextpoint1[1], arrtextpoint1[2]),
                            text = text.TextString
                        });
                    }
                }

                if (listLine.Count > 0)
                {
                    foreach (var line in listLine)
                    {
                        dynamic startpointarr = line.StartPoint;
                        dynamic endpointarr = line.EndPoint;

                        var startpoint = new XyzData((double)startpointarr[0], (double)startpointarr[1], 0);
                        var endpoint = new XyzData((double)endpointarr[0], (double)endpointarr[1], 0);

                        TextData textData = listpoint.MinBy2(x => x.point.ToXyzfit().DistanceTo(startpoint.ToXyz()));

                        _cadBeams.Add(new CadBeams()
                        {
                            StartPoint = startpoint,
                            EndPoint = endpoint,
                            Text = textData.text
                        });
                    }
                }

                GetBeamInfoCollection();
                dgBeam.ItemsSource = BeamInfoCollections;
            }

            ShowDialog();
        }

        public void GetBeamInfoCollection()
        {
            _beamInfoCollections.Clear();
            var dic = new Dictionary<string, List<BeamInfo>>();

            if (_cadBeams != null)
            {
                foreach (var cadbeam in _cadBeams)
                {
                    var text = cadbeam.Text;

                    var beamInfo = new BeamInfo(cadbeam.StartPoint.ToXyz(), cadbeam.EndPoint.ToXyz(), cadbeam.Text);
                    if (dic.ContainsKey(text))
                    {
                        dic[text].Add(beamInfo);
                    }
                    else
                    {
                        dic.Add(text, new List<BeamInfo> { beamInfo });
                    }
                }

                foreach (var pair in dic)
                {
                    var collection = new BeamInfoCollection
                    {
                        Text = pair.Key,
                        Width = pair.Value.Select(x => x.Width).FirstOrDefault(),
                        Height = pair.Value.Select(x => x.Height).FirstOrDefault(),
                        Number = pair.Value.Count
                    };
                    var b = pair.Value.ToList().Distinct(new BeamInfo.BeamInfoComparerByPoint()).ToList();
                    collection.BeamInfos = b;
                    _beamInfoCollections.Add(collection);
                }
            }

            BeamInfoCollections = new ObservableCollection<BeamInfoCollection>(_beamInfoCollections);
        }

        public void ModelBeams()
        {
            Hide();
            var selectedLevel = cbBeamLevels.SelectedItem as Level;
            try
            {
                _origin = AC.Selection.PickPoint();
            }
            catch (Exception e)
            {
            }

            if (_origin != null)
            {
                _origin = new XYZ(_origin.X, _origin.Y, 0);
                var max = BeamInfoCollections.Select(x => x.BeamInfos.Count).Sum();
                var progressView = new progressbar();
                progressView.Show();

                using var tg = new TransactionGroup(AC.Document, "Model Columns");
                tg.Start();
                foreach (var beamInfoCollection in BeamInfoCollections)
                {
                    if (progressView.Flag == false)
                    {
                        break;
                    }

                    var height = Convert.ToInt32(beamInfoCollection.Height);
                    var width = Convert.ToInt32(beamInfoCollection.Width);
                    var baseOffset =Double.Parse(txtBeamOffset.Text) ;

                    var elementType = beamInfoCollection.ElementType = GetElementType(width, height);

                    if (elementType == null)
                    {
                        continue;
                    }
                    DeleteWarningSuper waringsuper = new DeleteWarningSuper();

                    foreach (var beamInfo in beamInfoCollection.BeamInfos)
                    {
                        using var tx = new Transaction(AC.Document, "Modeling Beam From Cad");
                        tx.Start();

                        FailureHandlingOptions failOpt = tx.GetFailureHandlingOptions();

                        failOpt.SetFailuresPreprocessor(waringsuper);

                        tx.SetFailureHandlingOptions(failOpt);

                        var fs = beamInfoCollection.ElementType as FamilySymbol;
                        if (fs.IsActive == false)
                        {
                            fs.Activate();
                        }

                        var p1 = beamInfo.StartPoint.Add(_origin - CadBeamOrigin.ToXyz());
                        var p2 = beamInfo.EndPoint.Add(_origin - CadBeamOrigin.ToXyz());
                        var line = Line.CreateBound(p1, p2);
                        var a = line.GetEndPoint(0);
                        var b = line.GetEndPoint(1);

                        a = a.EditZ(selectedLevel.Elevation);
                        b = b.EditZ(selectedLevel.Elevation);

                        var l = Line.CreateBound(a, b);
                        try
                        {
                            var beam = AC.Document.Create.NewFamilyInstance(l, fs, selectedLevel,
                                StructuralType.Beam);

                            // var mark = beam.SetParameterValueByName(BuiltInParameter.ALL_MODEL_MARK, beamInfoCollection.Mark);
                            var startOffsetParam = beam.LookupParameter("Start Level Offset");
                            if (startOffsetParam is { IsReadOnly: false })
                            {
                                startOffsetParam.Set(baseOffset.MmToFoot());
                            }

                            var endOffsetParam = beam.LookupParameter("End Level Offset");
                            if (endOffsetParam is { IsReadOnly: false })
                            {
                                endOffsetParam.Set(baseOffset.MmToFoot());
                            }
                        }
                        catch
                        {
                        }

                        AC.Document.Regenerate();
                        progressView.Create(max, "BeamModel");

                        tx.Commit();
                    }
                }

                tg.Assimilate();
                progressView.Close();
            }

            ShowDialog();
        }

        //Lấy dữ liệu cho dầm để tạo
        private void GetDataDefaultBeam()
        {
            var families = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_StructuralFraming)
                .OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Select(x => x.Family)
                .Where(x => x.StructuralMaterialType != StructuralMaterialType.Steel).DistinctBy(x => x.Id)
                .OrderBy(x => x.Name).ToList();
            CbFamilyBeamTypes.ItemsSource = families;
            CbFamilyBeamTypes.SelectedItem = families.FirstOrDefault();

            var levels = new FilteredElementCollector(document).OfClass(typeof(Level)).Cast<Level>()
                .OrderBy(x => x.Elevation).ToList();

            cbBeamLevels.ItemsSource = levels;
            cbBeamLevels.SelectedItem = levels.FirstOrDefault();
        }

        #endregion

        #region Columns

        public void SelectFormCadColumn()
        {
            Hide();
            dynamic a = ComRunningObject.GetActiveObjectByProgId("AutoCaD.Application");
            a.Visible = true; // bật hiển thị
            a.WindowState = 3; // thường là Maximized (tùy enum COM)
            a.ActiveDocument.Activate(); // kích hoạt document

            dynamic doc
                = a.Documents.Application.ActiveDocument;

            string[] arr = null;
            try
            {
                var pointCad = doc.Utility.GetPoint(Type.Missing, "Select point: ");
                arr = ((IEnumerable)pointCad).Cast<object>()
                    .Select(x => x.ToString())
                    .ToArray();
            }
            catch (Exception e)
            {
            }

            if (arr != null)
            {
                double[] arr1 = new double[3];
                int i = 0;
                foreach (var item in arr)
                {
                    arr1[i] = Convert.ToDouble(item);
                    i++;
                }

                CadBeamOrigin = new XyzData(arr1[0], arr1[1], arr1[2]);
                var newset = doc.SelectionSets.Add(Guid.NewGuid().ToString());

                newset.SelectOnScreen();

                if (newset.Count <= 0)
                {
                }

                List<dynamic> listText = new List<dynamic>();

                List<dynamic> listPolyline = new List<dynamic>();

                foreach (dynamic s in newset)
                {
                    if (s.EntityName == "AcDbPolyline")
                    {
                        listPolyline.Add(s);
                    }

                    if (s.EntityName == "AcDbText")
                    {
                        listText.Add(s);
                    }
                }


                List<TextData> listpoint = new List<TextData>();
                if (listText.Count > 0)
                {
                    foreach (var text in listText)
                    {
                        string[] arrtextpoint = ((IEnumerable)text.InsertionPoint).Cast<object>()
                            .Select(x => x.ToString())
                            .ToArray();

                        double[] arrtextpoint1 = new double[3];
                        int k = 0;
                        foreach (var item in arrtextpoint)
                        {
                            arrtextpoint1[k] = Convert.ToDouble(item);
                            k++;
                        }

                        listpoint.Add(new TextData()
                        {
                            point = new XYZ(arrtextpoint1[0], arrtextpoint1[1], arrtextpoint1[2]),
                            text = text.TextString
                        });
                    }
                }
                else
                {
                    listpoint = null;
                }

                foreach (var polyline in listPolyline)
                {
                    dynamic c = polyline.Coordinates;
                    if (c.Length == 8)
                    {
                        dynamic pointArray1 = polyline.Coordinate[0];
                        dynamic pointArray2 = polyline.Coordinate[1];
                        dynamic pointArray3 = polyline.Coordinate[2];
                        dynamic pointArray4 = polyline.Coordinate[3];
                        var point1 = new XyzData((double)pointArray1[0], (double)pointArray1[1], 0);
                        var point2 = new XyzData((double)pointArray2[0], (double)pointArray2[1], 0);
                        var point3 = new XyzData((double)pointArray3[0], (double)pointArray3[1], 0);
                        var point4 = new XyzData((double)pointArray4[0], (double)pointArray4[1], 0);
                        if (listText.Count > 0)
                        {
                            TextData textData = listpoint.MinBy2(x => x.point.ToXyzfit().DistanceTo(point1.ToXyz()));

                            cadRectangles.Add(new CadRectangle()
                            {
                                P1 = point1,
                                P2 = point2,
                                P3 = point3,
                                P4 = point4,
                                Mask = textData.text
                            });
                        }
                        else
                        {
                            cadRectangles.Add(new CadRectangle()
                            {
                                P1 = point1,
                                P2 = point2,
                                P3 = point3,
                                P4 = point4,
                                Mask = ""
                            });
                        }
                    }
                }

                GetColumnInfoCollections();
            }

            ShowDialog();
        }

        public void GetColumnInfoCollections()
        {
            columnInfoCollections.Clear();

            var dic = new Dictionary<ColumnInfo, List<ColumnInfo>>();

            if (cadRectangles != null)
            {
                foreach (var cadRectangle in cadRectangles)
                {
                    var points = cadRectangle.Points.Select(x => x.ToXyz()).ToList();
                    if (points.Count == 4)
                    {
                        var columnInfo = new ColumnInfo(points, cadRectangle.Mask);
                        if (dic.ContainsKey(columnInfo))
                        {
                            dic[columnInfo].Add(columnInfo);
                        }
                        else
                        {
                            dic.Add(columnInfo, new List<ColumnInfo> { columnInfo });
                        }
                    }
                }

                foreach (var pair in dic)
                {
                    if ((pair.Key.Height / pair.Key.Width) < 5)
                    {
                        var collection = new ColumnInfoCollection
                        {
                            Width = pair.Key.Width,
                            Height = pair.Key.Height,
                            Number = pair.Value.Count,
                            Text = pair.Key.Text
                        };

                        var b = pair.Value.ToList().ToList();
                        collection.ColumnInfos = b;
                        columnInfoCollections.Add(collection);
                    }
                }
            }

            ColumnInfoCollections = new ObservableCollection<ColumnInfoCollection>(columnInfoCollections);
            dgColumn.ItemsSource = ColumnInfoCollections;
        }

        public void ModelColumn()
        {
            Hide();
            try
            {
                _origin = AC.Selection.PickPoint();
            }
            catch (Exception e)
            {
            }

            if (_origin != null)
            {
                try
                {
                    _origin = new XYZ(_origin.X, _origin.Y, 0);
                    var max = ColumnInfoCollections.Select(x => x.ColumnInfos.Count).Sum();
                    var progressView = new progressbar();
                    progressView.Show();
                    using (var tg = new TransactionGroup(AC.Document, "Model Columns"))
                    {
                        tg.Start();

                        DeleteWarningSuper waringsuper = new DeleteWarningSuper();

                        foreach (var columnInfoCollection in ColumnInfoCollections)
                        {
                            if (progressView.Flag == false)
                            {
                                break;
                            }

                            var width = Convert.ToInt32(columnInfoCollection.Width);
                            var height = Convert.ToInt32(columnInfoCollection.Height);


                            var elementType = columnInfoCollection.ElementType = GetColElementType(width, height);

                            if (elementType == null)
                            {
                                continue;
                            }

                            foreach (var columnInfo in columnInfoCollection.ColumnInfos)
                            {
                                using var tx = new Transaction(AC.Document, "Modeling Column From Cad");
                                tx.Start();
                                FailureHandlingOptions failOpt = tx.GetFailureHandlingOptions();

                                failOpt.SetFailuresPreprocessor(waringsuper);

                                tx.SetFailureHandlingOptions(failOpt);

                                var center = columnInfo.Center.Add(_origin - CadBeamOrigin.ToXyz());

                                var fs = columnInfoCollection.ElementType as FamilySymbol;
                                if (fs.IsActive == false)
                                {
                                    fs.Activate();
                                }

                                try
                                {
                                    var selectedBaseLevel = cbBaseLevel.SelectedItem as Level;
                                    var selectedTopLevel = cbTopLevel.SelectedItem as Level;
                                    var topOffset = double.Parse(txtTopOffset.Text);
                                    var baseOffset = double.Parse(txtBaseOffset.Text);

                                    var column = AC.Document.Create.NewFamilyInstance(center, fs,
                                        selectedBaseLevel, StructuralType.Column);

                                    tx.Commit();
                                    tx.Start();

                                    var pBaseLevel = column.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM);
                                    var pTopLevel = column.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM);

                                    if (pTopLevel != null && !pTopLevel.IsReadOnly)
                                        pTopLevel.Set(selectedTopLevel.Id);

                                    var topOffsetParam = column.LookupParameter("Top Offset");
                                    if (topOffsetParam == null)
                                    {
                                        topOffsetParam = column.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM);
                                    }
                                    if (topOffsetParam == null)
                                    {
                                        topOffsetParam = column.get_Parameter(BuiltInParameter.SCHEDULE_TOP_LEVEL_OFFSET_PARAM);
                                    }

                                    if (topOffsetParam != null && !topOffsetParam.IsReadOnly)
                                    {
                                        topOffsetParam.Set(topOffset.MmToFoot());
                                    }

                                    var baseOffsetParam = column.LookupParameter("Base Offset");
                                    if (baseOffsetParam == null)
                                    {
                                        baseOffsetParam = column.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM);
                                    }
                                    if (baseOffsetParam == null)
                                    {
                                        baseOffsetParam = column.get_Parameter(BuiltInParameter.SCHEDULE_BASE_LEVEL_OFFSET_PARAM);
                                    }

                                    if (baseOffsetParam != null && !baseOffsetParam.IsReadOnly)
                                    {
                                        baseOffsetParam.Set(baseOffset.MmToFoot());
                                    }

                                    var rotateAxis = center.CreateLineByPointAndDirection(XYZ.BasisZ);
                                    ElementTransformUtils.RotateElement(AC.Document, column.Id, rotateAxis,
                                        columnInfo.Rotation);
                                    progressView.Create(max, "ColumnModel");
                                }
                                catch
                                {
                                }

                                tx.Commit();
                            }
                        }

                        tg.Assimilate();
                        progressView.Close();
                    }

                }

                catch
                {
                }
            }
        }

        private ElementType GetColElementType(double width, double height)
        {
            ElementType elementType = null;

            using (var tx = new Transaction(AC.Document, "Duplicate Type"))
            {
                tx.Start();
                //update element in revit
                AC.Document.Regenerate();

                var selectedItem = cboColumnFamilyType.SelectedItem as Family;

                var columnTypes = selectedItem.GetFamilySymbolIds().Select(x =>document.GetElement(x) )
                    .Cast<FamilySymbol>().ToList();

                foreach (var familySymbol in columnTypes)
                {
                    var bParam = familySymbol.LookupParameter(cbWidthColumn.SelectedItem as string);
                    var hParam = familySymbol.LookupParameter(cbHeightCol.SelectedItem as string);
                    var bInMM = Convert.ToInt32(bParam.AsDouble().FootToMm());
                    var hInMM = Convert.ToInt32(hParam.AsDouble().FootToMm());

                    if (width == bInMM && height == hInMM)
                    {
                        elementType = familySymbol;
                    }
                }

                if (elementType == null)
                {
                    //Duplicate Column Type
                    var type = columnTypes.FirstOrDefault();

                    var newTypeName = "Column" + "_" + width + "x" + height;

                    if (columnTypes.Select(x => x.Name).Contains(newTypeName))
                    {
                        newTypeName = newTypeName + " Ignore existed name";
                    }

                    while (true)
                    {
                        try
                        {
                            elementType = type?.Duplicate(newTypeName);
                            break;
                        }
                        catch
                        {
                            newTypeName += ".";
                        }
                    }
                    if (elementType != null)
                    {
                        elementType.LookupParameter(cbWidthColumn.SelectedItem as string).Set(width.MmToFoot());
                        elementType.LookupParameter(cbHeightCol.SelectedItem as string).Set(height.MmToFoot());
                    }
                }

                tx.Commit();
            }
            return elementType;
        }

        #endregion

        #region Floor

        public void SelectFloorFromCad()
        {
            Hide();

            dynamic a = ComRunningObject.GetActiveObjectByProgId("AutoCaD.Application");
            a.Visible = true; // bật hiển thị
            a.WindowState = 3; // thường là Maximized (tùy enum COM)
            a.ActiveDocument.Activate(); // kích hoạt document

            dynamic doc
                = a.Documents.Application.ActiveDocument;

            string[] arrPoint = null;
            try
            {
                var pointCad = doc.Utility.GetPoint(Type.Missing, "Select point: ");
                arrPoint = ((IEnumerable)pointCad).Cast<object>()
                    .Select(x => x.ToString())
                    .ToArray();
            }
            catch (Exception e)
            {

            }

            if (arrPoint != null)
            {
                double[] arrPoint1 = new double[3];

                int i = 0;

                foreach (var item in arrPoint)
                {
                    arrPoint1[i] = Convert.ToDouble(item);
                    i++;
                }

                CadBeamOrigin = new XyzData(arrPoint1[0], arrPoint1[1], arrPoint1[2]);
                var newset = doc.SelectionSets.Add(Guid.NewGuid().ToString());

                newset.SelectOnScreen();

                List<dynamic> listText = new List<dynamic>();

                List<dynamic> listPolylines = new List<dynamic>();

                foreach (dynamic s in newset)
                {
                    if (s.EntityName == "AcDbPolyline")
                    {
                        listPolylines.Add(s);
                    }
                    if (s.EntityName == "AcDbText")
                    {
                        listText.Add(s);
                    }

                }

                foreach (var polyline in listPolylines)
                {
                    dynamic c = polyline.Coordinates;
                    var ct = Enumerable.Count(c) / 2;
                    var slabPoints = new List<XyzData>();
                    for (int j = 0; j < ct; j++)
                    {
                        dynamic pointarr = polyline.Coordinate[j];
                        var point = new XyzData((double)pointarr[0], (double)pointarr[1], 0);
                        slabPoints.Add(point);
                    }


                    for (int item = 0; item < slabPoints.Count; item++)
                    {

                        for (int item1 = 1; item1 < slabPoints.Count; item1++)
                        {
                            if (item < item1)
                            {
                                XYZ displacement = slabPoints[item].ToXyz().Subtract(slabPoints[item1].ToXyz());
                                double distance = displacement.GetLength();
                                if (distance < 0.08)
                                {
                                    slabPoints.RemoveAt(item1);
                                }
                            }
                        }
                    }

                    //MessageBox.Show(slabPoints.Count.ToString());

                    ListPoint.Add(slabPoints);

                    var collection = new FloorInfoCollection
                    {
                        Area = Math.Round(polyline.Area / 1000000, 1)
                    };

                    floorInfoCollections.Add(collection);
                }

                FloorInfoCollections = new ObservableCollection<FloorInfoCollection>(floorInfoCollections);
                dgFloor.ItemsSource = FloorInfoCollections;
            }

            ShowDialog();
        }
        public void ModelFloor()
        {
            Hide();
            try
            {
                _origin = AC.Selection.PickPoint();
            }
            catch (Exception e)
            {

            }

            if (_origin != null)
            {
                var max = FloorInfoCollections.Count;
                _origin = new XYZ(_origin.X, _origin.Y, 0);
                var progressView = new progressbar();
                progressView.Show();

                using var tg = new TransactionGroup(AC.Document, "Model Floor");
                tg.Start();

                DeleteWarningSuper waringsuper = new DeleteWarningSuper();

                foreach (var listpoint in ListPoint)
                {
                    if (progressView.Flag == false)
                    {
                        break;
                    }

                    using (var tx = new Transaction(AC.Document, "Modeling Column From Cad"))
                    {
                        tx.Start();
                        FailureHandlingOptions failOpt = tx.GetFailureHandlingOptions();

                        failOpt.SetFailuresPreprocessor(waringsuper);

                        tx.SetFailureHandlingOptions(failOpt);

                        CurveArray curvearr = new CurveArray();

                        var selectedLevel = cbFloorLevel.SelectedItem as Level;

                        for (int i = 0; i < listpoint.Count - 1; i++)
                        {
                            var p1 = listpoint[i].ToXyz().Add(_origin - CadBeamOrigin.ToXyz());
                            var p2 = listpoint[i + 1].ToXyz().Add(_origin - CadBeamOrigin.ToXyz());

                            p1 = p1.EditZ(selectedLevel.Elevation);
                            p2 = p2.EditZ(selectedLevel.Elevation);

                            var l = Line.CreateBound(p1, p2);

                            Curve curve = l;

                            curvearr.Append(curve);
                        }

                        var pe = listpoint[listpoint.Count - 1].ToXyz().Add(_origin - CadBeamOrigin.ToXyz()).EditZ(selectedLevel.Elevation);
                        var pt = listpoint[0].ToXyz().Add(_origin - CadBeamOrigin.ToXyz()).EditZ(selectedLevel.Elevation);

                        XYZ displacement = pt.Subtract(pe);
                        double distance = displacement.GetLength();

                        if (distance > 0.08)
                        {
                            Curve cv = Line.CreateBound(pe, pt);
                            curvearr.Append(cv);
                        }

                        try
                        {
                            var floorType = cbFloorType.SelectedItem as FloorType;
                            Floor floor;
                            var cl = new CurveLoop();
                            curvearr.ToCurves().ForEach(x => cl.Append(x));
                            floor = Floor.Create(AC.Document, new List<CurveLoop>() { cl }, floorType.Id, selectedLevel.Id);
                            var offsetParam = floor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM);

                            offsetParam.Set(double.Parse(txtFloorOffset.Text).MmToFoot());
                        }
                        catch
                        {

                        }
                        progressView.Create(max, "FloorModel");
                        tx.Commit();
                    }
                }

                tg.Assimilate();
                progressView.Close();
            }

        }
        private void GetDataDefaultFloor()
        {
            var families = new FilteredElementCollector(AC.Document).OfCategory(BuiltInCategory.OST_Floors)
                .OfClass(typeof(FloorType)).Cast<FloorType>().ToList();

            cbFloorType.ItemsSource = families;
            cbFloorType.SelectedItem = families.FirstOrDefault();

            var levels = new FilteredElementCollector(document).OfClass(typeof(Level)).Cast<Level>()
                .OrderBy(x => x.Elevation).ToList();

            cbFloorLevel.ItemsSource = levels;
            cbFloorLevel.SelectedItem = levels.FirstOrDefault();
        }

        #endregion
        private ElementType GetElementType(double width, double heigt)
        {
            ElementType elementType = null;

            using var tx = new Transaction(AC.Document, "Duplicate Type");
            tx.Start();
            //update element in revit
            AC.Document.Regenerate();

            var selectedItem = CbFamilyBeamTypes.SelectedItem as Family;

            var beamTypes = selectedItem.GetFamilySymbolIds().Select(x => document.GetElement(x))
                .Cast<FamilySymbol>().ToList();

            foreach (var familySymbol in beamTypes)
            {
                var bParameter = familySymbol.LookupParameter(cbBeamWidth.SelectedItem as string);

                var binMm = Convert.ToInt32(bParameter.AsDouble().FootToMm());

                var hParameter = familySymbol.LookupParameter(cbBeamHeight.SelectedItem as string);

                var hinMm = Convert.ToInt32(hParameter.AsDouble().FootToMm());

                if (heigt == hinMm && width == binMm)
                {
                    elementType = familySymbol;
                }
            }

            if (elementType == null)
            {
                //Duplicate Column Type
                var type = beamTypes.FirstOrDefault();

                var newTypeName = "Beams" + "" + width + "x" + heigt;

                var i = 1;
                if (beamTypes.Select(x => x.Name).Contains(newTypeName))
                {
                    newTypeName = $"{newTypeName} (1)";
                }

                while (true)
                {
                    try
                    {
                        elementType = type?.Duplicate(newTypeName);
                        break;
                    }
                    catch
                    {
                        i++;
                        newTypeName = $"{newTypeName} ({i})";
                    }
                }

                if (elementType != null)
                {
                    elementType.LookupParameter(cbBeamWidth.SelectedItem as string).Set(width.MmToFoot());
                    elementType.LookupParameter(cbBeamHeight.SelectedItem as string).Set(Utils.Utils.MmToFoot(heigt));
                }
            }

            tx.Commit();
            return elementType;
        }

        private void CbFamilyBeamTypes_OnSelected(object sender, RoutedEventArgs e)
        {
            var familySelected = CbFamilyBeamTypes.SelectedItem as Family;
            if (familySelected == null) return;

            var first = document.GetElement(familySelected.GetFamilySymbolIds().FirstOrDefault()) as FamilySymbol;
            if (first != null)
            {
                var data = first.GetOrderedParameters()
                    .Where(x => x.StorageType == StorageType.Double).Select(x => x.Definition.Name)
                    .ToList();
                cbBeamWidth.ItemsSource = data;

                cbBeamWidth.SelectedItem = data.FirstOrDefault();

                cbBeamHeight.ItemsSource = data;

                cbBeamHeight.SelectedItem = data.Skip(1).FirstOrDefault();
            }
        }

        private void CboColumnFamilyType_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var familySelected = cboColumnFamilyType.SelectedItem as Family;
            if (familySelected == null) return;

            var first = document.GetElement(familySelected.GetFamilySymbolIds().FirstOrDefault()) as FamilySymbol;
            if (first != null)
            {
                var data = first.GetOrderedParameters()
                    .Where(x => x.StorageType == StorageType.Double).Select(x => x.Definition.Name)
                    .ToList();
                cbWidthColumn.ItemsSource = data;

                cbWidthColumn.SelectedItem = data.FirstOrDefault();

                cbHeightCol.ItemsSource = data;

                cbHeightCol.SelectedItem = data.Skip(1).FirstOrDefault();
            }
        }
    }

    public class TextData
    {
        public XYZ point;
        public string text;
    }

    public class XyzData
    {
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }

        public XyzData(double x, double y, double z)
        {
            X = x;
            Y = y;
            Z = z;
        }

        public XyzData Mid(XyzData other)
        {
            var a = (X + other.X) / 2;
            var b = (Y + other.Y) / 2;
            var c = (Z + other.Z) / 2;
            return new XyzData(a, b, c);
        }
    }

    public class CadBeams
    {
        public XyzData StartPoint { get; set; }

        public XyzData EndPoint { get; set; }

        public string Text { get; set; }
    }

    public class CadRectangle
    {
        public XyzData P1 { get; set; }
        public XyzData P2 { get; set; }
        public XyzData P3 { get; set; }
        public XyzData P4 { get; set; }

        public string Mask;

        public List<XyzData> Points => new() { P1, P2, P3, P4 };
    }

    public class ColumnInfoCollection : ObservableObject
    {
        public List<ColumnInfo> ColumnInfos { get; set; } = new List<ColumnInfo>();

        private double _width;

        public double Width
        {
            get => _width;
            set
            {
                _width = value;
                OnPropertyChanged();
            }
        }

        private double _height;

        public double Height
        {
            get => _height;
            set
            {
                _height = value;
                OnPropertyChanged();
            }
        }

        private string text;

        public string Text
        {
            get => text;
            set
            {
                text = value;
                OnPropertyChanged();
            }
        }

        public ElementType ElementType { get; set; }

        public int Number { get; set; }


        public ColumnInfoCollection()
        {
        }
    }

    public class ColumnInfo
    {
        public XYZ Center { get; set; }
        public Line WidthLine { get; set; }
        public Line HeightLine { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public double Rotation { get; set; }


        public string Text { get; set; }

        public ColumnInfo(List<XYZ> points, string text)
        {
            GetInfo(points);
            Text = text;
        }

        public ColumnInfo()
        {
        }

        private void GetInfo(List<XYZ> points)
        {
            //Center
            Center = new XYZ(points.Average(x => x.X), points.Average(x => x.Y), points.Average(x => x.Z));
            var p1 = points[0];
            var p2 = points[1];
            var p3 = points[2];
            var p4 = points[3];
            var l1 = Line.CreateBound(p1, p2);
            var l2 = Line.CreateBound(p2, p3);
            if (l1.Length >= l2.Length)
            {
                HeightLine = l1;
                WidthLine = l2;
            }
            else
            {
                HeightLine = l2;
                WidthLine = l1;
            }

            var direction = HeightLine.Direction;

            if (direction.Y < 0)
            {
                direction = -direction;
            }

            Width = Math.Round(WidthLine.Length.FootToMm(), 1);

            Height = Math.Round(HeightLine.Length.FootToMm(), 1);

            Rotation = new XYZ(direction.X, direction.Y, 0).AngleTo(XYZ.BasisY);
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;
            if (GetType() != obj.GetType()) return false;
            return obj is ColumnInfo columnInfo;
        }

        public override int GetHashCode()
        {
            return 0;
        }
    }

    public class FloorInfoCollection : ObservableObject
    {

        private double _area;

        public double Area
        {
            get => _area;

            set
            {
                _area = value;
                OnPropertyChanged();
            }
        }
    }
}