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
            SketchLine line1 = (SketchLine) swLine1;
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
        /// https://www.bilibili.com/video/BV1C54y1W7d7?p=1
        /// </summary>
        /// <param name="swApp"></param>
        public void SolidWorksAcademy2019P1(SldWorks swApp)
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
            swFeat= swModel.FeatureByPositionReverse(0);
            swFeat.Name = "PipeSketch";
            swModel.Extension.SelectByID2("PipeSketch", "SKETCH", 0, 0, 0, false, 0, null, 0);
            swFeat = swModel.FeatureManager.FeatureExtrusion3(true, false, false,
                (int) swEndConditions_e.swEndCondBlind, 0, 0.8, 0, false, false, false, false, 0, 0, false, false,
                false, false, false, false, false, 0, 0, false);
            swFeat.Name = "PipeModel";
            swModel.ForceRebuild3(true);
            swModel.ViewZoomtofit2();
            Configuration swConfig = swModel.GetActiveConfiguration();
            CustomPropertyManager swCustPropMgr = swConfig.CustomPropertyManager;
            swCustPropMgr.Add3("Description", (int) swCustomInfoType_e.swCustomInfoText, "Pipe",
                (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
            swCustPropMgr.Add3("Dimensions", (int)swCustomInfoType_e.swCustomInfoText, "800",
                (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
        }

    }
}
