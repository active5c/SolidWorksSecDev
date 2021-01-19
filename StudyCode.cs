using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SolidWorksSecDev
{

    public class StudyCode
    {
        /// <summary>
        /// How to get the length of a belt feature?
        /// https://www.bilibili.com/video/BV1Lh41127Td?p=5
        /// </summary>
        /// <param name="swApp"></param>
        public void BlueByteP5(SldWorks swApp)
        {
            string beltPartPath = string.Empty;
            string beltFeatureName = "Belt1";
            decimal beltLength = 0m;

            ModelDoc2 swModel = swApp.ActiveDoc;//get the active doc
            if (swModel.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY) return;//check assemdoc
            //false to get the top-level features and all child features in the FeatureManager design tree
            //Return Value:Array of all of the features
            object[] featArr = swModel.FeatureManager.GetFeatures(false);//获取整个特征树中的所有特征
            //Traverse all features in the assem
            foreach (var item in featArr)
            {
                Feature swFeature = (Feature)item;
                //check if feature is sketch
                if (swFeature.GetTypeName2() == "ProfileFeature")//获取特征类型
                {

                    Feature swOwnerFeature = swFeature.GetOwnerFeature();//返回父特征


                    //check if owner feature is a belt type feature and its name is beltfeaturename
                    if (swOwnerFeature.GetTypeName2() == "Belt" && swOwnerFeature.Name == beltFeatureName)
                    {
                        Sketch swSketch = swFeature.GetSpecificFeature2();
                        SketchSegment[] swSketchSegmentArr = swSketch.GetSketchSegments();//获取草图中的部件
                        foreach (var sketchSegment in swSketchSegmentArr)
                        {
                            if (sketchSegment.ConstructionGeometry == false)
                                beltLength += (decimal)sketchSegment.GetLength();
                        }
                        Feature swCompFeature = swFeature.GetNextSubFeature();
                        if (swCompFeature != null)
                        {
                            Component2 swComponent = swCompFeature.GetSpecificFeature2();
                            beltPartPath = swComponent.GetPathName();
                        }
                    }
                }
            }
            ModelDoc2 swcompModelDoc = swApp.OpenDoc(beltPartPath, (int)swDocumentTypes_e.swDocPART);
            if (swcompModelDoc == null) return;
            swcompModelDoc.Visible = true;
            //添加自定义属性
            swcompModelDoc.Extension.CustomPropertyManager[""].Add2("Belt Length in mm",
                (int)swCustomInfoType_e.swCustomInfoText, Math.Round(beltLength, 2).ToString());

            swcompModelDoc.Save();
            swApp.QuitDoc(swcompModelDoc.GetTitle());
        }

        /// <summary>
        /// Can you split BOM?
        /// https://www.bilibili.com/video/BV1Lh41127Td?p=6
        /// </summary>
        /// <param name="swApp"></param>
        public void BlueByteP6(SldWorks swApp)
        {
            //precondition:Bill of materials seleted
            //前置条件：Drawing中必须先选中BOM表
            ModelDoc2 swModel = swApp.ActiveDoc;
            if (swModel.GetType() != (int)swDocumentTypes_e.swDocDRAWING) return;//check DRAWING
            SelectionMgr swSelectionMgr = swModel.SelectionManager;
            TableAnnotation parentBomTableAnnotation = swSelectionMgr.GetSelectedObject6(1, 0);
            //切割BOM表，从第二行往后切割
            TableAnnotation newBomTableAnnotation = parentBomTableAnnotation.Split((int)swTableSplitLocations_e.swTableSplit_AfterRow, 1);
            Annotation newBomAnnotation = newBomTableAnnotation.GetAnnotation();
            //获取新的表格的坐标位置
            double[] position = newBomAnnotation.GetPosition();
            //clear the selection 
            swModel.ClearSelection();
            //select new Bom,选中切割后的表，根据坐标
            swModel.Extension.SelectByID2("", "ANNOTATIONTABLES", position[0], position[1], position[2], false, 0, null, 0);
            swModel.EditCut();//将选中的注释，剪切
            DrawingDoc swDrawingDoc = (DrawingDoc)swModel;
            swDrawingDoc.ActivateSheet("Sheet1");//激活第一张表
            swModel.Paste();//粘贴
        }

        /// <summary>
        /// Feature selection inside an assembly
        /// 
        /// </summary>
        /// <param name="swApp"></param>
        public void BlueByteP8(SldWorks swApp)
        {
            //get active model in the solidworks sessin
            ModelDoc2 swModel = swApp.ActiveDoc;
            AssemblyDoc swAssy = (AssemblyDoc)swModel;
            //get an array of the top level components
            var swComps = swAssy.GetComponents(true);
            Component2 swComp = swComps[0];
            //get the first feature of swComp
            Feature swFeat = swComp.FirstFeature();
            string selectionName = String.Empty;
            //traverse first level
            while (swFeat != null)
            {
                if (swFeat.GetTypeName2() == "RefPlane" && swFeat.Name == "Right Plane")
                {
                    string featType = String.Empty;
                    selectionName = swFeat.GetNameForSelection(out featType);
                    Debug.Print(selectionName);
                }
                swFeat = swFeat.GetNextFeature();
            }
            //select right plane
            swModel.Extension.SelectByID2(selectionName, "PLANE", 0, 0, 0, false, 0, null, 0);
        }

        /// <summary>
        /// delete the note of tool box in drawing
        /// https://www.bilibili.com/video/BV19K411M7Y7?p=3
        /// </summary>
        /// <param name="swApp"></param>
        public void CADSharpP3(SldWorks swApp)
        {
            //前提是打开工程图，然后选中视图
            //代码的作用是删除标准件的标号
            ModelDoc2 swModel = swApp.ActiveDoc;
            SelectionMgr swSelectionMgr = swModel.SelectionManager;
            View swView = swSelectionMgr.GetSelectedObject6(1, -1);//-1 means ignore mark
            if (swView == null) return;
            //get the selected drawing view(arry)
            List<Annotation> swAnnotations = swView.GetAnnotations();
            //traverse the drawing view for all notes
            foreach (var item in swAnnotations)
            {
                //item.Select3(true, null);选中，作为阶段性调试观察
                //get the entities the note is attached to
                if (item.GetType() != (int)swAnnotationType_e.swNote) return;
                object[] swAttachedEntities = item.GetAttachedEntities3();
                if (swAttachedEntities.Length == 0) return;
                Entity swEntity = (Entity)swAttachedEntities[0];//assume only one entity
                //swEntity.Select4(true, null);选中，作为阶段性调试观察
                //get the underlying component for the entity
                Component2 swComponent = swEntity.GetComponent();
                //determine if component is from the tool box，if yes, delete it
                Debug.Print(swComponent.GetPathName());
                //判断路径中是否包含"solidworks data"//通常tool box存档在该目录中
                if (!swComponent.GetPathName().ToLower().Contains("solidworks data")) return;
                item.Select3(true, null);//选中annotations
                swModel.EditDelete();//删除
            }
        }

        /// <summary>
        /// Creat a left hand version of a part,打开目录中的模型，创建配置，镜像零件
        /// https://www.bilibili.com/video/BV19K411M7Y7?p=4
        /// </summary>
        /// <param name="swApp"></param>
        public void CADSharpP4(SldWorks swApp)
        {
            //Creat a left hand version of a part
            //Precondition
            //SolidWork id open
            //DIR exists and contains parts that meet the following specifications:
            //part can be mirrored about the front plane
            //single solid body
            //the x-y plane is called "Front Plane"
            //one config called "Default", and has "Suppress new features" enabled

            //get the active document
            //ModelDoc2 swModel = swApp.ActiveDoc;
            string dir = @"D:\model\";
            var files = Directory.EnumerateFiles(dir, "*.sldprt", SearchOption.AllDirectories);
            if (files == null) return;
            foreach (var item in files)
            {
                int errors = 0;
                int warnings = 0;
                ModelDoc2 swModel = swApp.OpenDoc6(item, (int)swDocumentTypes_e.swDocPART,
                    (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);
                //add another configuration
                swModel.AddConfiguration3("Opposite", "", "",
                    (int)swConfigurationOptions2_e.swConfigOption_DoDisolveInBOM);
                //select the front plane and mark it
                //SelectByID2适合于选中特征树中有名字name对象,注意mark参数
                swModel.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 2, null, 0);
                //select the body and mark it，get all the bodys of the part,assume only one body
                PartDoc swPart = (PartDoc)swModel;
                object[] swBodies = swPart.GetBodies2((int)swBodyType_e.swAllBodies, false);
                Body2 swBody = (Body2)swBodies[0];
                SelectionMgr swSelMgr = (SelectionMgr)swModel.SelectionManager;
                SelectData swSelData = (SelectData)swSelMgr.CreateSelectData();
                swSelData.Mark = 256;
                swBody.Select2(true, swSelData);
                //mirror the body,镜像实体
                swModel.FeatureManager.InsertMirrorFeature2(true, false, false, false, 0);
                //delete the original body，删除实体
                swBody.Select2(false, null);
                swModel.FeatureManager.InsertDeleteBody();
                //save the part
                swModel.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref errors, ref warnings);
                swApp.CloseDoc(swModel.GetPathName());
            }
            swApp.SendMsgToUser("Done!");
        }

        /// <summary>
        /// Cylinder Length and Diameter Program//新建零件绘制一个圆柱形
        /// https://www.bilibili.com/video/BV1M5411n7GH
        /// </summary>
        /// <param name="swApp"></param>
        public void Mein3d(SldWorks swApp)
        {
            //Add a "Hello World!" pop-up box to a macro
            //swApp.SendMsgToUser("Hello World!");
            //understand the concept of:1.Type casting and Early binding
            //获取用户默认模版
            string defaultPartTemplate =
                swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);

            //Early binding,早期绑定
            //create a part document
            ModelDoc2 swModel = swApp.NewDocument(defaultPartTemplate, 0, 0, 0);
            if (swModel == null) return;
            //Type casting,类型转换
            PartDoc swPartDoc = (PartDoc)swModel;

            //SketchManager,FeatureManager,Creat first subroutine
            //Open a sketch
            swModel.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swModel.SketchManager.InsertSketch(true);
            //create variables for the length and diameter of the cylinder
            double cylinderLength = 120d / 1000d;
            double cylinderDiameter = 20d / 1000d;
            //draw a circle，画圆
            swModel.SketchManager.CreateCircle(0, 0, 0, 0, cylinderDiameter / 2d, 0);
            //extrude the cylinder,拉伸
            swModel.FeatureManager.FeatureExtrusion3(true, false, false,
                (int)swEndConditions_e.swEndCondBlind,
                (int)swEndConditions_e.swEndCondBlind,
                cylinderLength, 0, false, false, false, false,
                0, 0, false, false, false, false, true, true, true,
                (int)swStartConditions_e.swStartSketchPlane, 0, false);
        }

        /// <summary>
        /// swept feature，扫掠特征的创建
        /// https://www.bilibili.com/video/BV1vK411M7Fi
        /// </summary>
        /// <param name="swApp"></param>
        public void MirzaCenanovic(SldWorks swApp)
        {
            //swApp.CloseAllDocuments(true);//DANGEROUS很危险，确保已经保存
            //新建零件
            ModelDoc2 swModel = swApp.NewPart();
            //PartDoc swPartDoc = (PartDoc)swModel;
            //create 3D sketch，创建3D草图
            swModel.Insert3DSketch2(false);
            //create lines，创建两条线
            SketchSegment swLine1 = swModel.CreateLine2(0, 0, 0, 0, 0, 100d / 1000d);
            SketchSegment swLine2 = swModel.CreateLine2(0, 0, 100d / 1000d, 0, 100d / 1000d, 100d / 1000d);
            //create fillet，创建圆角
            swLine1.Select(false);
            swLine2.Select(true);
            swModel.SketchFillet2(10d / 1000d, 0);
            swModel.Insert3DSketch2(true);

            //create sketch plane，根据点和法线创建平面
            SketchLine line1 = (SketchLine)swLine1;
            SketchPoint swPoint1 = line1.GetStartPoint2();
            swLine1.Select(false);
            swPoint1.Select(true);
            swModel.CreatePlanePerCurveAndPassPoint3(true, true);

            //create profile sketch，创建扫掠截面草图
            swModel.Extension.SelectByID2("Plane1", "PLANE", 0, 0, 0, false, 0, null, 0);
            swModel.SelectionManager.GetSelectedObject6(1, -1);
            swModel.BlankRefGeom();//hide the plane，隐藏创建的平面
            swModel.SketchManager.InsertSketch(false);
            swModel.SketchManager.CreateCircle(0, 0, 0, 5d / 1000d, 0, 0);
            swModel.SketchManager.CreateCircle(0, 0, 0, 7d / 1000d, 0, 0);
            swModel.SketchManager.InsertSketch(true);

            //swept feature，选中截面和引导线，创建扫掠特征
            swModel.Extension.SelectByID2("Sketch1", "SKETCH", 0, 0, 0, false, 0, null, 0);
            swModel.Extension.SelectByID2("3DSketch1", "SKETCH", 0, 0, 0, true, 0, null, 0);
            swModel.FeatureManager.InsertProtrusionSwept3(false, false, 0, true, false,
                 0, 0, false, 0, 0, 0, 0, true, false, false, 0, false);

        }

        /// <summary>
        /// 创建一个拉伸零件，然后添加自定义属性
        /// https://www.bilibili.com/video/BV1C54y1W7d7?p=2
        /// </summary>
        /// <param name="swApp"></param>
        public void SolidWorksAcademy2019P2(SldWorks swApp)
        {
            //新建零件
            ModelDoc2 swModel = swApp.NewPart();
            //重命名三个基准面，这个选择的方法是按照倒序选择的，0位是原点
            Feature swFeat = swModel.FeatureByPositionReverse(3);
            swFeat.Name = "Front";
            swFeat = swModel.FeatureByPositionReverse(2);
            swFeat.Name = "Top";
            swFeat = swModel.FeatureByPositionReverse(1);
            swFeat.Name = "Right";
            //绘制两个圆
            swModel.Extension.SelectByID2("Top", "PLANE", 0, 0, 0, false, 0, null, 0);
            swModel.InsertSketch2(true);
            swModel.CreateCircleByRadius2(0, 0, 0, 0.5);
            swModel.CreateCircleByRadius2(0, 0, 0, 0.2);
            swModel.InsertSketch2(true);
            //
            swFeat = swModel.FeatureByPositionReverse(0);
            swFeat.Name = "PipeSketch";
            swModel.Extension.SelectByID2("PipeSketch", "SKETCH", 0, 0, 0, false, 0, null, 0);
            swFeat = swModel.FeatureManager.FeatureExtrusion3(true, false, false,
                (int)swEndConditions_e.swEndCondBlind, 0, 0.8, 0, false, false, false, false, 0, 0, false, false,
                false, false, false, false, false, 0, 0, false);
            swFeat.Name = "PipeModel";
            swModel.ForceRebuild3(true);
            swModel.ViewZoomtofit2();
            Configuration swConfig = swModel.GetActiveConfiguration();
            CustomPropertyManager swCustPropMgr = swConfig.CustomPropertyManager;
            swCustPropMgr.Add3("Description", (int)swCustomInfoType_e.swCustomInfoText, "Pipe",
                (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
            swCustPropMgr.Add3("Dimensions", (int)swCustomInfoType_e.swCustomInfoText, "800",
                (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
        }

        /// <summary>
        /// 绘制一个钣金零件，打孔，绘制圆角，定义装配面
        /// https://www.bilibili.com/video/BV1C54y1W7d7?p=5
        /// </summary>
        public void SolidWorksAcademyP5CreateAnglePart(SldWorks swApp)
        {
            //定义参数,可以从窗体中获取参数，或者其他数据源
            double xLength = 100d / 1000d;
            double yLength = 200d / 1000d;
            double width = 125d / 1000d;
            double thickness = 8d / 1000d;

            double boltHole = 7.5d / 1000d;
            double pipeHole = 25.5d / 1000d;
            double rad1 = 15d / 1000d;
            double rad2 = 20d / 1000d;
            double x1 = 25d / 1000d;
            double x2 = 25d / 1000d;

            string mateRef1 = "Ref1";
            string mateRefHole = "RefHole";

            string targetFolder = @"E:\Videos\SolidWorks Secondary Development\WorksAcademyP4";
            string fileName = "AnglePart";


            //获取用户默认模版
            string defaultPartTemplate =
                swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);

            ModelDoc2 swModel = swApp.NewDocument(defaultPartTemplate, 0, 0, 0);
            if (swModel == null) return;

            Feature swFeat = swModel.FeatureByPositionReverse(3);
            swFeat.Name = "Front";
            swModel.Extension.SelectByID2("Front", "PLANE", 0, 0, 0, false, 0, null, 0);

            swModel.InsertSketch2(true);
            swModel.CreateLine2(0, 0, 0, xLength, 0, 0);
            //swModel.AddDimension2(0, 0, 0);
            swModel.CreateLine2(0, 0, 0, 0, yLength, 0);
            //swModel.AddDimension2(0, 0, 0);

            //int markHorizontal = 2;
            //int markVertical = 4;
            //swModel.Extension.SelectByID2("Point1<Origin", "EXTSKETCHSEGMENT", 0, 0, 0, false, markHorizontal | markVertical, null, 0);            
            //swModel.SketchManager.FullyDefineSketch(true, true,
            //    (int)swSketchFullyDefineRelationType_e.swSketchFullyDefineRelationType_Vertical |
            //    (int)swSketchFullyDefineRelationType_e.swSketchFullyDefineRelationType_Horizontal, true,
            //    (int)swAutodimScheme_e.swAutodimSchemeBaseline, null,
            //    (int)swAutodimScheme_e.swAutodimSchemeBaseline, null,
            //    (int)swAutodimHorizontalPlacement_e.swAutodimHorizontalPlacementBelow,
            //    (int)swAutodimVerticalPlacement_e.swAutodimVerticalPlacementLeft);

            object datumDisp = "Point1@Origin";
            swModel.SketchManager.FullyDefineSketch(true, true,
                (int)swSketchFullyDefineRelationType_e.swSketchFullyDefineRelationType_Vertical |
                (int)swSketchFullyDefineRelationType_e.swSketchFullyDefineRelationType_Horizontal, true,
                (int)swAutodimScheme_e.swAutodimSchemeBaseline, datumDisp,
                (int)swAutodimScheme_e.swAutodimSchemeBaseline, datumDisp,
                (int)swAutodimHorizontalPlacement_e.swAutodimHorizontalPlacementBelow,
                (int)swAutodimVerticalPlacement_e.swAutodimVerticalPlacementLeft);

            swModel.InsertSketch2(true);

            swFeat = swModel.FeatureByPositionReverse(0);
            swFeat.Name = "Sketch1";
            swModel.Extension.SelectByID2("Sketch1", "SKETCH", 0, 0, 0, false, 0, null, 0);
            //基体法兰
            swModel.FeatureManager.InsertSheetMetalBaseFlange2(thickness, false, thickness, width, 0, true, 0, 0, 0, null, false, 0, 0, 0, 0, false, false, false, false);

            //change entity name
            swModel.Extension.SelectByID2("", "FACE", thickness, yLength / 2, -width / 2, false, 0, null, 0);
            SelectionMgr swSelMgr = swModel.SelectionManager;
            Face2 swFace = swSelMgr.GetSelectedObject6(1, -1);
            PartDoc swPartDoc = (PartDoc)swModel;
            swPartDoc.SetEntityName(swFace, mateRef1);

            swModel.Extension.SelectByID2("", "FACE", xLength / 2, thickness, -width / 2, false, 0, null, 0);
            swModel.InsertSketch2(true);
            swModel.CreateCircleByRadius2(xLength - x1, x2, 0, boltHole);
            swModel.CreateCircleByRadius2(xLength - x1, width - x2, 0, boltHole);
            swModel.InsertSketch2(true);
            //拉伸切除
            swModel.FeatureManager.FeatureCut3(true, false, false,
                (int)swEndConditions_e.swEndCondThroughAll,
                (int)swEndConditions_e.swEndCondBlind, 0, 0, false, false, false, false, 0, 0,
                false, false, false, false, false, true, true, false, false, false,
                (int)swEndConditions_e.swEndCondMidPlane, 0, true);

            Entity swEntity = swPartDoc.GetEntityByName(mateRef1, (int)swSelectType_e.swSelFACES);
            swEntity.Select4(false, null);
            swModel.InsertSketch2(true);
            swModel.CreateCircleByRadius2(width / 2, yLength - (pipeHole * 2), 0, pipeHole);
            swModel.AddDiameterDimension(0, 0, 0);
            swModel.InsertSketch2(true);
            //拉伸切除
            swModel.FeatureManager.FeatureCut3(true, false, false,
                (int)swEndConditions_e.swEndCondThroughAll,
                (int)swEndConditions_e.swEndCondBlind, 0, 0, false, false, false, false, 0, 0,
                false, false, false, false, false, true, true, false, false, false,
                (int)swEndConditions_e.swEndCondMidPlane, 0, true);

            //change entity name
            swModel.Extension.SelectByID2("", "FACE", thickness / 2, yLength - pipeHole, -width / 2, false, 0, null, 0);
            swFace = swSelMgr.GetSelectedObject6(1, -1);
            swPartDoc.SetEntityName(swFace, mateRefHole);

            //圆角特征
            swModel.Extension.SelectByID2("", "EDGE", xLength, thickness / 2, 0, false, 0, null, 0);
            swModel.Extension.SelectByID2("", "EDGE", xLength, thickness / 2, -width, true, 0, null, 0);
            swModel.FeatureManager.FeatureFillet3((int)swFeatureFilletOptions_e.swFeatureFilletUniformRadius, rad2,
                0, 0, 0, 0, 0, null, null, null, null, null, null, null);
            swModel.Extension.SelectByID2("", "EDGE", thickness / 2, yLength, 0, false, 0, null, 0);
            swModel.Extension.SelectByID2("", "EDGE", thickness / 2, yLength, -width, true, 0, null, 0);
            swModel.FeatureManager.FeatureFillet3((int)swFeatureFilletOptions_e.swFeatureFilletUniformRadius, rad1,
                0, 0, 0, 0, 0, null, null, null, null, null, null, null);

            swModel.ClearSelection2(true);
            swModel.ViewZoomtofit2();

            string root = targetFolder;
            if (!Directory.Exists(root)) Directory.CreateDirectory(root);

            swModel.SaveAs3(targetFolder + "\\" + fileName + ".sldprt", (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                (int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);
            swApp.CloseDoc(fileName + ".sldprt");
        }

        /// <summary>
        /// 绘制一个管件，定义装配面
        /// https://www.bilibili.com/video/BV1C54y1W7d7?p=5
        /// </summary>
        public void SolidWorksAcademyP5CreatePipe(SldWorks swApp)
        {
            //定义参数,可以从窗体中获取参数，或者其他数据源
            double outsideDiameter = 50d / 1000d;
            double insideDiameter = 40d / 1000d;
            double length = 50d / 1000d;

            string mateOutsideFace = "PipeOutsideFace";
            string mateBase = "PipeFace";

            string targetFolder = @"E:\Videos\SolidWorks Secondary Development\WorksAcademyP4";
            string fileName = "Pipe";

            //获取用户默认模版
            string defaultPartTemplate =
                swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);

            ModelDoc2 swModel = swApp.NewDocument(defaultPartTemplate, 0, 0, 0);
            if (swModel == null) return;

            Feature swFeat = swModel.FeatureByPositionReverse(3);
            swFeat.Name = "Front";
            swModel.Extension.SelectByID2("Front", "PLANE", 0, 0, 0, false, 0, null, 0);

            swModel.InsertSketch2(true);
            swModel.CreateCircleByRadius2(0, 0, 0, insideDiameter / 2d);
            swModel.CreateCircleByRadius2(0, 0, 0, outsideDiameter / 2d);
            swModel.InsertSketch2(true);

            swFeat = swModel.FeatureManager.FeatureExtrusion3(true, false, false, 0, 0, length, 0, false, false, false,
                false, 0, 0, false, false, false, false, false, false, false, 0, 0, false);

            //change entity name
            swModel.Extension.SelectByID2("", "FACE", outsideDiameter / 2, 0, length / 2, false, 0, null, 0);
            SelectionMgr swSelMgr = swModel.SelectionManager;
            Face2 swFace = swSelMgr.GetSelectedObject6(1, -1);
            PartDoc swPartDoc = (PartDoc)swModel;
            swPartDoc.SetEntityName(swFace, mateOutsideFace);
            //change entity name
            swModel.Extension.SelectByID2("", "FACE", insideDiameter / 2 + (outsideDiameter - insideDiameter) / 4, 0, 0, false, 0, null, 0);
            swFace = swSelMgr.GetSelectedObject6(1, -1);
            swPartDoc.SetEntityName(swFace, mateBase);

            swModel.ClearSelection2(true);
            swModel.ViewZoomtofit2();

            string root = targetFolder;
            if (!Directory.Exists(root)) Directory.CreateDirectory(root);

            swModel.SaveAs3(targetFolder + "\\" + fileName + ".sldprt", (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                (int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);
            swApp.CloseDoc(fileName + ".sldprt");

        }

        /// <summary>
        /// 将钣金和管件装配在一起
        /// https://www.bilibili.com/video/BV1C54y1W7d7?p=6
        /// </summary>
        public void SolidWorksAcademyP6(SldWorks swApp)
        {
            //调用上面两个方法，创建两个零件
            SolidWorksAcademyP5CreateAnglePart(swApp);
            SolidWorksAcademyP5CreatePipe(swApp);

            //开始装配
            string targetFolder = @"E:\Videos\SolidWorks Secondary Development\WorksAcademyP4";
            string fileName = "TestAssembly";
            string pipeFileName = "Pipe";
            string anglePartFileName = "AnglePart";

            //获取用户默认模版
            string defaultAssemblyTemplate =
                swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateAssembly);

            ModelDoc2 swModel = swApp.NewDocument(defaultAssemblyTemplate, 0, 0, 0);
            if (swModel == null) return;
            AssemblyDoc swAssy = (AssemblyDoc) swModel;

            string[] xCompNames=new string[2];
            xCompNames[0] = targetFolder+"\\"+pipeFileName+".sldprt";
            xCompNames[1] = targetFolder + "\\" + anglePartFileName + ".sldprt";
            object compNames = xCompNames;

            string[] xCoorSysNames=new string[2];
            xCoorSysNames[0] = "Coordinate System1";
            xCoorSysNames[1] = "Coordinate System1";
            object coorSysName = xCoorSysNames;
            
            double pipleLength = 50d / 1000d;
            double angelPartyLength = 200d / 1000d;
            var tMatrix=new double[]
            {
                0,0,1,
                0,1,0,
                -1,0,0,
                pipleLength,angelPartyLength,0,
                0,0,0,0,
                1,0,0,
                0,1,0,
                0,0,1,
                0,0,0,
                0,0,0,0
            };
            object transformationMatrix = tMatrix;

            swAssy.AddComponents3(compNames, transformationMatrix, coorSysName);

            string comp = anglePartFileName + "-1@" + fileName;
            swModel.Extension.SelectByID2(comp, "COMPONENT", 0, 0, 0, false, 0, null, 0);
            swAssy.FixComponent();

            string mateRef1 = "Ref1";
            string mateRefHole = "RefHole";
            string mateOutsideFace = "PipeOutsideFace";
            string mateBase = "PipeFace";

            Component2 swComp = swAssy.GetComponentByName(anglePartFileName + "-1");
            ModelDoc2 swCompModel = swComp.GetModelDoc2();
            PartDoc swCompPart = (PartDoc) swCompModel;
            Entity swEntity = swCompPart.GetEntityByName(mateRefHole, (int) swSelectType_e.swSelFACES);
            Entity swFace1 = swComp.GetCorrespondingEntity(swEntity);
            swEntity = swCompPart.GetEntityByName(mateRef1, (int)swSelectType_e.swSelFACES);
            Entity swFace3 = swComp.GetCorrespondingEntity(swEntity);

            swComp = swAssy.GetComponentByName(pipeFileName + "-1");
            swCompModel = swComp.GetModelDoc2();
            swCompPart = (PartDoc) swCompModel;
            swEntity = swCompPart.GetEntityByName(mateOutsideFace, (int) swSelectType_e.swSelFACES);
            Entity swFace2 = swComp.GetCorrespondingEntity(swEntity);
            swEntity = swCompPart.GetEntityByName(mateBase, (int)swSelectType_e.swSelFACES);
            Entity swFace4 = swComp.GetCorrespondingEntity(swEntity);
            //同心配合
            int errorCode;
            swFace1.Select4(false, null);
            swFace2.Select4(true, null);
            swAssy.AddMate3((int) swMateType_e.swMateCONCENTRIC, (int) swMateAlign_e.swMateAlignALIGNED, false, 0, 0, 0,0, 0, 0, 0, 0, false, out errorCode);
            //距离配合
            double thickness = 8d / 1000d;
            swFace3.Select4(false, null);
            swFace4.Select4(true, null);
            swAssy.AddMate3((int)swMateType_e.swMateDISTANCE, (int)swMateAlign_e.swMateAlignALIGNED, false, thickness, thickness, thickness, 0, 0, 0, 0, 0, false, out errorCode);

            swModel.ViewZoomtofit2();
            swModel.ForceRebuild3(false);
            string root = targetFolder;
            if (!Directory.Exists(root)) Directory.CreateDirectory(root);

            swModel.SaveAs3(targetFolder + "\\" + fileName + ".sldasm", (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                (int)swSaveAsOptions_e.swSaveAsOptions_CopyAndOpen);
            swApp.CloseDoc(fileName + ".sldasm");
            Process.Start("explorer.exe", targetFolder);
        }

        /// <summary>
        /// 
        /// https://www.bilibili.com/video/BV1oX4y1N74h?p=1
        /// </summary>
        /// <param name="swApp"></param>
        public void CADCoderP1(SldWorks swApp)
        {
            
        }
    }
}
