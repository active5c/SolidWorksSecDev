using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data.SqlTypes;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using View = SolidWorks.Interop.sldworks.View;

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
        /// https://www.bilibili.com/video/BV1Lh41127Td?p=8
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

        #region CADSharp2

        /// <summary>
        /// https://www.bilibili.com/video/BV1Mp4y1Y7Bd?p=14
        /// Opening, Saving, and Exporting Documents
        /// </summary>
        /// <param name="swApp"></param>
        public void CADSharp2P14(SldWorks swApp)
        {
            //string defaultPartTemplate = swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            //ModelDoc2 swModel1 = swApp.NewDocument(defaultPartTemplate, 0, 0, 0);
            ModelDoc2 swModel2 = swApp.OpenDoc6(@"E:\Videos\SolidWorks Secondary Development\SWModel\Part1.SLDPRT", (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0);
            DocumentSpecification swDocSpec = swApp.GetOpenDocSpec(@"E:\Videos\SolidWorks Secondary Development\SWModel\Assem4.SLDASM");
            swDocSpec.DisplayState = "Transparent";
            swDocSpec.UseLightWeightDefault = false;
            swDocSpec.LightWeight = true;
            ModelDoc2 swModel3 = swApp.OpenDoc7(swDocSpec);

            ModelDoc2 swModel = swApp.ActivateDoc2("Part1.SLDPRT", true, 0);
            //另存为eDrawing版本
            string newPath = swModel.GetPathName().Replace("SLDPRT", "eprt");
            swModel.Extension.SaveAs(newPath, 0, 1 + 2, null, 0, 0);
            swApp.CloseDoc(swModel.GetPathName());
        }

        /// <summary>
        /// https://www.bilibili.com/video/BV1Mp4y1Y7Bd?p=15
        /// Working With Configurations
        /// </summary>
        /// <param name="swApp"></param>
        public void CADSharp2P15(SldWorks swApp)
        {
            ModelDoc2 swModel = swApp.ActiveDoc;
            //Test for the number of configuration
            if (swModel.GetConfigurationCount() == 1)
            {
                Debug.Print("Only one configuration");
                //Add three configs
                swModel.AddConfiguration3("A", "", "", 0);
                swModel.AddConfiguration3("B", "", "", 0);
                swModel.AddConfiguration3("C", "", "", 0);
            }
            else
            {
                Debug.Print("More than one configuration exists.");
            }
            //get the active configuration's name
            ConfigurationManager swConfigMgr = swModel.ConfigurationManager;
            Configuration swConfig = swConfigMgr.ActiveConfiguration;
            string strConfig = swConfig.Name;
            Debug.Print(strConfig);
            //delete the first configuration
            string[] vConfigNames = swModel.GetConfigurationNames();
            swModel.DeleteConfiguration2(vConfigNames[0]);
            //Cycle through configs and change names
            vConfigNames = swModel.GetConfigurationNames();
            for (int i = 0; i < vConfigNames.Length; i++)
            {
                //swModel.EditConfiguration3(vConfigNames[i], "Config" + (i + 1), "", "", 0);
                swConfig = swModel.GetConfigurationByName(vConfigNames[i]);
                swConfig.Name = "Config" + (i + 1);
            }
            //Activate the first configuration
            vConfigNames = swModel.GetConfigurationNames();
            swModel.ShowConfiguration2(vConfigNames[0]);
        }

        /// <summary>
        /// https://www.bilibili.com/video/BV1Mp4y1Y7Bd?p=16
        /// Working With Custom Properties
        /// </summary>
        /// <param name="swApp"></param>
        public void CADSharp2P16(SldWorks swApp)
        {
            ModelDoc2 swModel = swApp.ActiveDoc;
            //Access cust prop at doc level
            CustomPropertyManager swCustPropMgr = swModel.Extension.CustomPropertyManager[""];
            swCustPropMgr.Add2("Description", (int)swCustomInfoType_e.swCustomInfoText, "your desc");
            swCustPropMgr.Set("Author", "your name");
            string strValue;
            string strResolved;
            swCustPropMgr.Get3("PartNo", true, out strValue, out strResolved);
            Debug.Print(strValue + "   " + strResolved);
            swCustPropMgr.Delete("PartNo");

            //Access cust prop at config level
            string[] vConfigNames = swModel.GetConfigurationNames();
            for (int i = 0; i < vConfigNames.Length; i++)
            {
                swCustPropMgr = swModel.Extension.CustomPropertyManager[vConfigNames[i]];
                swCustPropMgr.Add2("PartNo", (int)swCustomInfoType_e.swCustomInfoText, strValue + "-" + vConfigNames[i]);
                swCustPropMgr.Add2("Material", (int)swCustomInfoType_e.swCustomInfoText, (char)34 + "SW-Material@" + swModel.GetTitle() + (char)34);
            }
        }

        /// <summary>
        /// https://www.bilibili.com/video/BV1Mp4y1Y7Bd?p=17
        /// Selection
        /// </summary>
        /// <param name="swApp"></param>
        public void CADSharp2P17(SldWorks swApp)
        {
            ModelDoc2 swModel = swApp.ActiveDoc;
            SelectionMgr swSelMgr = swModel.SelectionManager;
            Debug.Print(swSelMgr.GetSelectedObjectType3(-1, 1).ToString());
            for (int i = 0; i < swSelMgr.GetSelectedObjectCount2(-1); i++)
            {
                if (swSelMgr.GetSelectedObjectType3(i, -1) == (int)swSelectType_e.swSelBODYFEATURES)
                {
                    Feature swFeat = swSelMgr.GetSelectedObject6(i, -1);
                    MessageBox.Show(swFeat.Name);
                }
                else
                {
                    MessageBox.Show("Please select a feature from the feature manager tree.");
                }
            }

            if (swSelMgr.GetSelectedObjectType3(1, -1) == (int)swSelectType_e.swSelEDGES)
            {
                Edge swEdge = swSelMgr.GetSelectedObject6(1, -1);
                swModel.ClearSelection2(true);//清除选择
                //选择两个相接面
                Face2[] vFace = swEdge.GetTwoAdjacentFaces2();
                for (int i = 0; i < vFace.Length; i++)
                {
                    Face2 swFace = vFace[i];
                    Entity swEnt = (Entity)swFace;
                    SelectData swSelData = default(SelectData);
                    swEnt.Select4(true, swSelData);
                    Debug.Print(swSelData.X + "," + swSelData.Y + "," + swSelData.Z); ;
                }
            }
            else
            {
                MessageBox.Show("Please select an edge.");
            }

            //SelectByID2
            //根据空间坐标点选择面，线
            swModel.Extension.SelectByID2("", "FACE", -0.04, -0.140, 0, false, 0, null, 0);
            swModel.Extension.SelectByID2("", "EDGE", -0.05, -0.140, 0, true, 0, null, 0);
            //根据特征树名称选择草图点，面，特征
            swModel.Extension.SelectByID2("Point1@Origin", "EXTSKETCHPOINT", -50, -150, 0, true, 0, null, 0);
            swModel.Extension.SelectByID2("Front Plane", "PLANE", -50, -150, 0, true, 0, null, 0);
            swModel.Extension.SelectByID2("Cut-Extrudel", "BODYFEATURE", -50, -150, 0, true, 0, null, 0);

        }

        /// <summary>
        /// https://www.bilibili.com/video/BV1Mp4y1Y7Bd?p=18
        /// System and Document Options
        /// </summary>
        /// <param name="swApp"></param>
        public void CADSharp2P18(SldWorks swApp)
        {
            //Input Dimension Value,false取消勾选，true勾选
            if (swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate))
            {
                swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);
            }
            else
            {
                swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, true);
            }
            Debug.Print(swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate).ToString());
            //API搜索：System Options and Document Properties
            //数值类型的设置
            if (swApp.GetUserPreferenceDoubleValue((int)swUserPreferenceDoubleValue_e.swViewTransitionHideShowComponent) == 0)
            {
                swApp.SetUserPreferenceDoubleValue((int)swUserPreferenceDoubleValue_e.swViewTransitionHideShowComponent, 0.5);
            }
            else
            {
                swApp.SetUserPreferenceDoubleValue((int)swUserPreferenceDoubleValue_e.swViewTransitionHideShowComponent, 0.5);
            }
            Debug.Print(swApp.GetUserPreferenceDoubleValue((int)swUserPreferenceDoubleValue_e.swViewTransitionHideShowComponent).ToString());
            //Document Properties
            //标注保持两位小数点
            ModelDoc2 swModel = swApp.ActiveDoc;
            swModel.Extension.SetUserPreferenceInteger(
                (int)swUserPreferenceIntegerValue_e.swDetailingLinearDimPrecision,
                (int)swUserPreferenceOption_e.swDetailingDimension, 2);
        }

        /// <summary>
        /// https://www.bilibili.com/video/BV1Mp4y1Y7Bd?p=19
        /// Working With Sketches
        /// </summary>
        /// <param name="swApp"></param>
        public void CADSharp2P19(SldWorks swApp)
        {
            string defaultPartTemplate = swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            ModelDoc2 swModel = swApp.NewDocument(defaultPartTemplate, 0, 0, 0);

            //ModelDoc2 swModel = swApp.ActiveDoc;
            SketchManager swSkethMgr = swModel.SketchManager;
            //select front plane and insert sketch
            swModel.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swSkethMgr.InsertSketch(true);
            //turn on direct addition to database
            swSkethMgr.AddToDB = true;
            //create sketch entities
            SketchSegment swSketchSeg = swSkethMgr.CreateLine(0, 0, 0, 0, 0.05, 0);
            //添加约束
            swModel.SketchAddConstraints("sgVERTICAL2D");
            //将添加尺寸时出现的弹窗关闭（取消勾选）
            if (swApp.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate)) swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);
            //添加尺寸
            swModel.AddVerticalDimension2(-0.01, 0.025, 0);

            //make first sketch pt coincident with origin
            //first select both pts then add ralation
            //添加直线起点与原点重合配合
            swModel.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, false, 0, null, 0);
            swModel.Extension.SelectByID2("", "SKETCHPOINT", 0, 0, 0, true, 0, null, 0);
            swModel.SketchAddConstraints("sgCOINCIDENT");

            swSkethMgr.CreateTangentArc(0, 0.05, 0, 0.05, 0.05, 0, (int)swTangentArcTypes_e.swForward);
            //smart dimension
            swModel.AddDimension2(0.025, 0.08, 0);

            swSkethMgr.CreateLine(0.05, 0.05, 0, 0, 0, 0);
            swModel.AddHorizontalDimension2(0.05, 0.025, 0);

            swSkethMgr.CreateCircleByRadius(0.025, 0.05, 0, 0.005);
            swModel.AddDimension2(0.025, 0.065, 0);

            //fully define sketch，草图完全定义
            //swModel.Extension.SelectByID2("", "EXSKETCHPOINT", 0, 0, 0, false, 6, null, 0);
            //swSkethMgr.FullyDefineSketch(true, true, 1023, true, 1, null, 1, null, 0, 0);

            //将添加尺寸时出现的弹窗打开（勾选）
            swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, true);
            //turn off direct addition to database
            swSkethMgr.AddToDB = false;
            swSkethMgr.InsertSketch(true);
            //swModel.ClearSelection2(true);
            //修改已经存在的草图
            //make sure one sketch is selected
            //swModel.Extension.SelectByID2("", "SKETCH", 0.05, 0.05, 0, false, 0, null, 0);
            SelectionMgr swSelMgr = swModel.SelectionManager;
            if (swSelMgr.GetSelectedObjectType3(1, -1) != (int)swSelectType_e.swSelSKETCHES || swSelMgr.GetSelectedObjectCount2(-1) > 1)
            {
                swModel.Extension.ShowSmartMessage("Please select a single sketch from the FeatureManager tree", 3000, false, true);
            }
            //open the selected sketch
            swSkethMgr.InsertSketch(false);
            //get the sketch
            Sketch swSketch = swSkethMgr.ActiveSketch;
            //get the sketch segments
            int intSketchEntCount = 0;
            SketchSegment[] vSketchSegments = swSketch.GetSketchSegments();
            for (int i = 0; i < vSketchSegments.Length; i++)
            {
                swSketchSeg = vSketchSegments[i];
                //for some reason this won't recognize true...
                if (!swSketchSeg.ConstructionGeometry) intSketchEntCount++;
            }
            //close the sketch
            swSkethMgr.InsertSketch(true);
            //display the result
            swModel.Extension.ShowSmartMessage("Sketch entities:" + intSketchEntCount, 3000, false, true);
        }

        /// <summary>
        /// https://www.bilibili.com/video/BV1Mp4y1Y7Bd?p=20
        /// Working with Features Part A
        /// </summary>
        /// <param name="swApp"></param>
        public void CADSharp2P20(SldWorks swApp)
        {
            string defaultPartTemplate = swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            ModelDoc2 swModel = swApp.NewDocument(defaultPartTemplate, 0, 0, 0);
            SketchManager swSketchMgr = swModel.SketchManager;
            //create plate sketch profile
            swModel.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swSketchMgr.InsertSketch(true);
            swSketchMgr.CreateCircleByRadius(0, 0, 0, 0.05);
            swModel.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, false, 6, null, 0);
            swSketchMgr.FullyDefineSketch(true, true, 1023, true, 1, null, 1, null, 0, 0);
            FeatureManager swFeatureMgr = swModel.FeatureManager;
            //create mid-plane base extrusion
            Feature swFeat = swFeatureMgr.FeatureExtrusion2(true, false, false, 6, 0, 0.01, 0, false, false, false, false, 0, 0, false, false, false, false, true, true, true, 0, 0, false);
            //create extrude cut profile
            swModel.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swSketchMgr.InsertSketch(true);
            swSketchMgr.CreateCircleByRadius(-0.04, 0, 0, 0.004);
            swModel.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, false, 6, null, 0);
            swSketchMgr.FullyDefineSketch(true, true, 1023, true, 1, null, 1, null, 0, 0);
            //create tow-side cut extrude
            swFeat = swFeatureMgr.FeatureCut3(false, false, false, 1, 1, 0, 0, false, false, false, false, 0, 0, false, false, false, false, false, true, true, true, true, false, 0, 0, false);
            //create axis 
            swModel.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swModel.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, true, 0, null, 0);
            swModel.InsertAxis2(true);
            //create circular pattern of extruded cut
            swModel.Extension.SelectByID2("Axis1", "AXIS", 0, 0, 0, false, 1, null, 0);
            swModel.Extension.SelectByID2("Cut-Extrude1", "BODYFEATURE", 0, 0, 0, true, 4, null, 0);
            swFeatureMgr.FeatureCircularPattern2(10, 2 * 4 * Math.Atan(1) / 10, false, "", false);
            //create fillet
            PartDoc swPart = (PartDoc)swModel;
            var vBodies = swPart.GetBodies2((int)swBodyType_e.swSolidBody, false);
            //assumes only one solid body in part
            Body2 swBody = (Body2)vBodies[0];
            var vFace = swBody.GetFaces();
            Face2 swFinalFace = default(Face2);
            double dblArea = 0;
            for (int i = 0; i < vFace.Length; i++)
            {
                Face2 swFace = (Face2)vFace[i];
                Surface swSurf = swFace.GetSurface();
                if (swSurf.IsCylinder())
                {
                    //拿到面积最大的圆柱面
                    if (swFace.GetArea() > dblArea)
                    {
                        dblArea = swFace.GetArea();
                        swFinalFace = swFace;
                    }
                }
            }
            Entity swEnt = (Entity)swFinalFace;
            swEnt.Select4(false, null);
            swFeatureMgr.FeatureFillet(2, 0.001, 0, 0, 0, 0, 0);

            //change the material
            swPart.SetMaterialPropertyName2("", "", "Plain Carbon Steel");
        }

        /// <summary>
        /// https://www.bilibili.com/video/BV1Mp4y1Y7Bd?p=21
        /// Working with Features Part B
        /// </summary>
        /// <param name="swApp"></param>
        public void CADSharp2P21(SldWorks swApp)
        {
            ModelDoc2 swModel = swApp.ActiveDoc;
            PartDoc swPart = (PartDoc)swModel;
            //修改拉伸深度
            Feature swFeat = swPart.FeatureByName("Boss-Extrude1");
            ExtrudeFeatureData2 swExtrudeFeatData = swFeat.GetDefinition();
            swExtrudeFeatData.SetDepth(true, swExtrudeFeatData.GetDepth(true) * 1.5);
            swFeat.ModifyDefinition(swExtrudeFeatData, swModel, null);
            //modify fillet 修改圆角特征,将圆角特征应用到所有的边线
            swFeat = swPart.FeatureByName("Fillet1");
            SimpleFilletFeatureData2 swFilletFeatData = swFeat.GetDefinition();
            swFilletFeatData.AccessSelections(swModel, null);
            var vBodies = swPart.GetBodies2((int)swBodyType_e.swSolidBody, false);
            //assumes only one solid body in part
            Body2 swBody = (Body2)vBodies[0];
            swFilletFeatData.Edges = swBody.GetEdges();
            swFeat.ModifyDefinition(swFilletFeatData, swModel, null);
        }

        /// <summary>
        /// https://www.bilibili.com/video/BV1Mp4y1Y7Bd?p=22
        /// Working with Features Part C
        /// </summary>
        /// <param name="swApp"></param>
        public void CADSharp2P22_1(SldWorks swApp)
        {
            ModelDoc2 swModel = swApp.ActiveDoc;
            Debug.Print(swModel.GetPathName());
            Feature swFeat = swModel.FirstFeature();
            while (swFeat != null)
            {
                Debug.Print("FeatName:" + swFeat.Name);
                Debug.Print("   Type1:" + swFeat.GetTypeName());
                Debug.Print("   Type2:" + swFeat.GetTypeName2());
                Feature swSubFeat = (Feature)swFeat.GetFirstSubFeature();
                while (swSubFeat != null)
                {
                    Debug.Print("   SubFeatName:" + swSubFeat.Name);
                    Debug.Print("       Type1:" + swSubFeat.GetTypeName());
                    Debug.Print("       Type2:" + swSubFeat.GetTypeName2());
                    swSubFeat = (Feature)swSubFeat.GetNextSubFeature();
                }
                swFeat = (Feature)swFeat.GetNextFeature();
            }
        }

        /// <summary>
        /// https://www.bilibili.com/video/BV1Mp4y1Y7Bd?p=22
        /// Working with Features Part C
        /// </summary>
        /// <param name="swApp"></param>
        public void CADSharp2P22_2(SldWorks swApp)
        {
            //修改P20中的代码
            string defaultPartTemplate = swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            ModelDoc2 swModel = swApp.NewDocument(defaultPartTemplate, 0, 0, 0);
            //create plate sketch profile
            //--------replaces SelectByID2--------
            Feature swFeat = swModel.FirstFeature();
            do
            {
                if (swFeat.GetTypeName() == "RefPlane") break;
                swFeat = swFeat.GetNextFeature();
            } while (swFeat != null);
            swFeat.Select2(false, 0);
            //-------------------------------------
            SketchManager swSketchMgr = swModel.SketchManager;
            swSketchMgr.InsertSketch(true);
            swSketchMgr.CreateCircleByRadius(0, 0, 0, 0.05);
            swModel.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, false, 6, null, 0);
            swSketchMgr.FullyDefineSketch(true, true, 1023, true, 1, null, 1, null, 0, 0);
            FeatureManager swFeatureMgr = swModel.FeatureManager;
            //create mid-plane base extrusion
            swFeat = swFeatureMgr.FeatureExtrusion2(true, false, false, 6, 0, 0.01, 0, false, false, false, false, 0, 0, false, false, false, false, true, true, true, 0, 0, false);
            //create extrude cut profile
            //--------replaces SelectByID2--------
            swFeat = swModel.FirstFeature();
            do
            {
                if (swFeat.GetTypeName() == "RefPlane") break;
                swFeat = swFeat.GetNextFeature();
            } while (swFeat != null);
            swFeat.Select2(false, 0);
            //-------------------------------------
            swSketchMgr.InsertSketch(true);
            swSketchMgr.CreateCircleByRadius(-0.04, 0, 0, 0.004);
            swModel.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, false, 6, null, 0);
            swSketchMgr.FullyDefineSketch(true, true, 1023, true, 1, null, 1, null, 0, 0);
            //create tow-side cut extrude
            swFeat = swFeatureMgr.FeatureCut3(false, false, false, 1, 1, 0, 0, false, false, false, false, 0, 0, false, false, false, false, false, true, true, true, true, false, 0, 0, false);
            //create axis 
            //--------replaces SelectByID2--------
            swFeat = swModel.FirstFeature();
            int j = 0;
            do
            {
                if (swFeat.GetTypeName() == "RefPlane")
                {
                    if (j == 1) break;
                    j++;
                }
                swFeat = swFeat.GetNextFeature();
            } while (swFeat != null);
            swFeat.Select2(false, 0);
            //-------------------------------------
            //--------replaces SelectByID2--------
            swFeat = swModel.FirstFeature();
            j = 0;
            do
            {
                if (swFeat.GetTypeName() == "RefPlane")
                {
                    if (j == 2) break;
                    j++;
                }
                swFeat = swFeat.GetNextFeature();
            } while (swFeat != null);
            swFeat.Select2(true, 0);
            //-------------------------------------
            swModel.InsertAxis2(true);
            //create circular pattern of extruded cut
            //--------replaces SelectByID2--------
            swFeat = swModel.FirstFeature();
            do
            {
                if (swFeat.GetTypeName() == "RefAxis") break;
                swFeat = swFeat.GetNextFeature();
            } while (swFeat != null);
            swFeat.Select2(false, 1);
            //-------------------------------------
            //--------replaces SelectByID2--------
            swFeat = swModel.FirstFeature();
            do
            {
                if (swFeat.GetTypeName() == "Cut") break;
                swFeat = swFeat.GetNextFeature();
            } while (swFeat != null);
            swFeat.Select2(true, 4);
            //-------------------------------------
            swFeatureMgr.FeatureCircularPattern2(10, 2 * 4 * Math.Atan(1) / 10, false, "", false);
            //create fillet
            PartDoc swPart = (PartDoc)swModel;
            var vBodies = swPart.GetBodies2((int)swBodyType_e.swSolidBody, false);
            //assumes only one solid body in part
            Body2 swBody = (Body2)vBodies[0];
            var vFace = swBody.GetFaces();
            Face2 swFinalFace = default(Face2);
            double dblArea = 0;
            for (int i = 0; i < vFace.Length; i++)
            {
                Face2 swFace = (Face2)vFace[i];
                Surface swSurf = swFace.GetSurface();
                if (swSurf.IsCylinder())
                {
                    //拿到面积最大的圆柱面
                    if (swFace.GetArea() > dblArea)
                    {
                        dblArea = swFace.GetArea();
                        swFinalFace = swFace;
                    }
                }
            }
            Entity swEnt = (Entity)swFinalFace;
            swEnt.Select4(false, null);
            swFeatureMgr.FeatureFillet(2, 0.001, 0, 0, 0, 0, 0);

            //change the material
            swPart.SetMaterialPropertyName2("", "", "Plain Carbon Steel");
        }


        /// <summary>
        /// https://www.bilibili.com/video/BV1Mp4y1Y7Bd?p=23
        /// Geometry and Topology Objects(几何与拓扑)
        /// </summary>
        /// <param name="swApp"></param>
        public void CADSharp2P23_1(SldWorks swApp)
        {
            ModelDoc2 swModel = swApp.ActiveDoc;
            PartDoc swPart = (PartDoc)swModel;
            var vBodies = swPart.GetBodies2((int)swBodyType_e.swSolidBody, false);
            //assumes only one solid body in part
            Body2 swBody = (Body2)vBodies[0];
            var vFace = swBody.GetFaces();
            Face2 swFinalFace = default(Face2);
            double dblArea = 0;
            for (int i = 0; i < vFace.Length; i++)
            {
                Face2 swFace = (Face2)vFace[i];
                Surface swSurf = swFace.GetSurface();
                if (swSurf.IsCylinder())
                {
                    //拿到面积最大的圆柱面
                    if (swFace.GetArea() > dblArea)
                    {
                        dblArea = swFace.GetArea();
                        swFinalFace = swFace;
                    }
                }
            }
            Entity swEnt = (Entity)swFinalFace;
            swEnt.Select4(false, null);
            FeatureManager swFeatureMgr = swModel.FeatureManager;
            swFeatureMgr.FeatureFillet(2, 0.001, 0, 0, 0, 0, 0);
        }

        /// <summary>
        /// https://www.bilibili.com/video/BV1Mp4y1Y7Bd?p=23
        /// Geometry and Topology Objects(几何与拓扑)
        /// </summary>
        /// <param name="swApp"></param>
        public void CADSharp2P23_2(SldWorks swApp)
        {
            ModelDoc2 swModel = swApp.ActiveDoc;
            SelectionMgr swSelMgr = swModel.SelectionManager;
            double dblTotalLength = 0;
            for (int i = 0; i < swSelMgr.GetSelectedObjectCount2(-1); i++)
            {
                if (swSelMgr.GetSelectedObjectType3(i, -1) == (int)swSelectType_e.swSelEDGES)
                {
                    Edge swEdge = swSelMgr.GetSelectedObject6(i, -1);
                    Curve swCurve = swEdge.GetCurve();
                    CurveParamData swCurveParams = swEdge.GetCurveParams3();
                    double dblLength = swCurve.GetLength3(swCurveParams.UMinValue, swCurveParams.UMaxValue);
                    dblTotalLength += dblLength;
                    //print to immediate window
                    Debug.Print("---EDGE #" + i);
                    switch (swCurve.Identity())
                    {
                        case 3001:
                            Debug.Print("   Type:line");
                            break;
                        case 3002:
                            Debug.Print("   Type:circle");
                            break;
                        case 3003:
                            Debug.Print("   Type:ellipse");
                            break;
                        case 3004:
                            Debug.Print("   Type:intersection");
                            break;
                        case 3005:
                            Debug.Print("   Type:b-curve");
                            break;
                        case 3006:
                            Debug.Print("   Type:sp-curve");
                            break;
                        case 3008:
                            Debug.Print("   Type:constant parameter");
                            break;
                        case 3009:
                            Debug.Print("   Type:trimmed");
                            break;
                    }
                    dblLength = Math.Round(dblLength, 4);
                    Debug.Print("   Length:" + dblLength * 1000d + "mm");
                }
            }
            dblTotalLength = Math.Round(dblTotalLength, 4);
            Debug.Print("---TotalLength:" + dblTotalLength * 1000d + "mm");
        }

        /// <summary>
        /// https://www.bilibili.com/video/BV1Mp4y1Y7Bd?p=24
        /// Creating A New Assembly
        /// </summary>
        /// <param name="swApp"></param>
        public void CADSharp2P24_1(SldWorks swApp)
        {
            ModelDoc2 swModel = swApp.ActiveDoc;
            if (swModel == null || swModel.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY) return;
            AssemblyDoc swAssy = (AssemblyDoc)swModel;
            SelectionMgr swSelMgr = swModel.SelectionManager;
            if (swSelMgr.GetSelectedObjectType3(1, -1) != (int)swSelectType_e.swSelFACES)
            {
                MessageBox.Show("Please select a face.");
                return;
            }
            
            //get the drawer face 
            Face2 swFace = swSelMgr.GetSelectedObject6(1, -1);
            Entity swDrawerFace = (Entity)swFace;
            swDrawerFace = swDrawerFace.GetSafeEntity();

            Collection<Entity> collDrawerEdge = new Collection<Entity>();
            Collection<Entity> collCompFace = new Collection<Entity>();
            GetDrawerEdge();

            for (int i = 0; i < collDrawerEdge.Count; i++)
            {
                Component2 swComp= AddComps();
                GetCompFace(swComp);
                AddMates(swComp, collCompFace[i],collDrawerEdge[i]);
            }
            swModel.ClearSelection2(true);
            swModel.ForceRebuild3(false);

            //-----------------------------------------------
            //局部方法1 获取圆
            void GetDrawerEdge()
            {
                //get the full circle edge(s) from drawer face
                var vEdges = swFace.GetEdges();
                Entity swDrawerEdge = default(Entity);
                for (int i = 0; i < vEdges.Length; i++)
                {
                    Edge swEdge = (Edge)vEdges[i];
                    Curve swCurve = swEdge.GetCurve();
                    if (swCurve.Identity() == (int)swCurveTypes_e.CIRCLE_TYPE) //3002,circle
                    {
                        Vertex swVertex = swEdge.GetStartVertex();
                        if (swVertex == null)
                        {
                            swDrawerEdge = (Entity)swEdge;
                            swDrawerEdge = swDrawerEdge.GetSafeEntity();
                            collDrawerEdge.Add(swDrawerEdge);
                        }
                    }
                }
                if (swDrawerEdge == null) return;
            }


            //局部方法2 插入子装配
            Component2 AddComps()
            {
                //add the component
                string strCompPath = @"E:\Videos\SolidWorks Secondary Development\SWModel\CADSharp2P20.SLDPRT";
                swApp.DocumentVisible(false, (int)swDocumentTypes_e.swDocPART);
                swApp.OpenDoc6(strCompPath, (int)swDocumentTypes_e.swDocPART, 0, "", 0, 0);
                Component2 swComp = swAssy.AddComponent4(strCompPath, "", 0, 0, 0);//XYZ边界区域的中心bounding box centre  
                swApp.DocumentVisible(true, (int)swDocumentTypes_e.swDocPART);
                return swComp;
            }

            //局部方法3 获取圆柱面
            void GetCompFace(Component2 swComp)
            {
                //get the cylindrical face(s) from knob(s)
                var vBodies = swComp.GetBodies3((int)swBodyType_e.swSolidBody, out _);
                Body2 swBody = (Body2)vBodies[0];
                var vFaces = swBody.GetFaces();
                Entity swCompFace = default(Entity);
                for (int i = 0; i < vFaces.Length; i++)
                {
                    swFace = (Face2)vFaces[i];
                    Surface swSurf = swFace.GetSurface();
                    if (swSurf.IsCylinder())
                    {
                        swCompFace = (Entity)swFace;
                        collCompFace.Add(swCompFace);
                        break;
                    }
                }
            }

            //局部方法4  添加配合
            void AddMates(Component2 swComp,Entity swCompFace,Entity swDrawerEdge)
            {
                //add mates
                //1.add coincident mate
                swDrawerFace.Select4(false, null);
                Debug.Print(swComp.Name2);
                //注意对"Front Plane@" + swComp.Name2 + "@Assem1"做修改，Front Plane是swComp中的面，Assem1是swAssy的名称
                swModel.Extension.SelectByID2("Front Plane@" + swComp.Name2 + "@Assem1", "PLANE", 0, 0, 0, true, 0, null, 0);
                swAssy.AddMate3((int)swMateType_e.swMateCOINCIDENT, (int)swMateAlign_e.swMateAlignALIGNED, false, 0, 0, 0, 0, 0, 0, 0, 0, false, out _);

                //2.add concentric mate
                swCompFace.Select4(false, null);
                swDrawerEdge.Select4(true, null);
                swAssy.AddMate3((int)swMateType_e.swMateCONCENTRIC, (int)swMateAlign_e.swMateAlignANTI_ALIGNED, false, 0, 0, 0, 0, 0, 0, 0, 0, false, out _);
            }

        }

        /// <summary>
        /// https://www.bilibili.com/video/BV1Mp4y1Y7Bd?p=24
        /// Creating A New Assembly
        /// </summary>
        /// <param name="swApp"></param>
        public void CADSharp2P24_2(SldWorks swApp)
        {
            ModelDoc2 swModel = swApp.ActiveDoc;
            AssemblyDoc swAssy = (AssemblyDoc) swModel;
            
            //get the arm face by searching components for face called "ArmFace"
            Entity swArmFace=GetArmFace();
            swArmFace = swArmFace.GetSafeEntity();
            //Get the full circular edge on the arm
            GetArmEdge();
            //Insety the handle
            Component2 swComp =AddComp();
            //Get a cylindrical face on the handle
            Entity swHandleCylFace = default(Entity);
            swHandleCylFace = GetHandleCylFace(swComp);
            //Get the required planar face on the handle
            GetHandlePlanarFace();
            //Add the Mate
            AddMate();

            swModel.ClearSelection2(true);
            swModel.ForceRebuild3(false);

            //----------局部方法-------------
            Entity GetArmFace()
            {
                return default(Entity);
            }

            Entity GetArmEdge()
            {
                return default(Entity);
            }

            Component2 AddComp()
            {
                return default(Component2);
            }

            Entity GetHandleCylFace(Component2 swHandleComp)
            {
                var vComps = swAssy.GetComponents(false);
                for (int i = 0; i < vComps.Length; i++)
                {
                    var vBodies = swHandleComp.GetBodies3((int) swBodyType_e.swSolidBody, out _);
                    Body2 swBody = (Body2) vBodies[0];
                    var vFace = swBody.GetFaces();
                    for (int j = 0; j < vFace.Length; j++)
                    {
                        Face2 swFace = vFace[i];
                        Entity swHandleCyl = (Entity) swFace;
                        if (swModel.GetEntityName(swHandleCylFace) == "HandleCylFace") return swHandleCyl;
                    }
                }
                return default(Entity);
            }

            void GetHandlePlanarFace()
            {
            }

            void AddMate()
            {
                
            }

        }

        /// <summary>
        /// https://www.bilibili.com/video/BV1Mp4y1Y7Bd?p=25
        /// Working With Existing Assemblies
        /// </summary>
        /// <param name="swApp"></param>
        public void CADSharp2P25(SldWorks swApp)
        {

        }

        /// <summary>
        /// https://www.bilibili.com/video/BV1Mp4y1Y7Bd?p=26
        /// Component Transforms
        /// </summary>
        /// <param name="swApp"></param>
        public void CADSharp2P26(SldWorks swApp)
        {

        }



        #endregion

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
            AssemblyDoc swAssy = (AssemblyDoc)swModel;

            string[] xCompNames = new string[2];
            xCompNames[0] = targetFolder + "\\" + pipeFileName + ".sldprt";
            xCompNames[1] = targetFolder + "\\" + anglePartFileName + ".sldprt";
            object compNames = xCompNames;

            string[] xCoorSysNames = new string[2];
            xCoorSysNames[0] = "Coordinate System1";
            xCoorSysNames[1] = "Coordinate System1";
            object coorSysName = xCoorSysNames;

            double pipleLength = 50d / 1000d;
            double angelPartyLength = 200d / 1000d;
            var tMatrix = new double[]
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
            PartDoc swCompPart = (PartDoc)swCompModel;
            Entity swEntity = swCompPart.GetEntityByName(mateRefHole, (int)swSelectType_e.swSelFACES);
            Entity swFace1 = swComp.GetCorrespondingEntity(swEntity);
            swEntity = swCompPart.GetEntityByName(mateRef1, (int)swSelectType_e.swSelFACES);
            Entity swFace3 = swComp.GetCorrespondingEntity(swEntity);

            swComp = swAssy.GetComponentByName(pipeFileName + "-1");
            swCompModel = swComp.GetModelDoc2();
            swCompPart = (PartDoc)swCompModel;
            swEntity = swCompPart.GetEntityByName(mateOutsideFace, (int)swSelectType_e.swSelFACES);
            Entity swFace2 = swComp.GetCorrespondingEntity(swEntity);
            swEntity = swCompPart.GetEntityByName(mateBase, (int)swSelectType_e.swSelFACES);
            Entity swFace4 = swComp.GetCorrespondingEntity(swEntity);
            //同心配合
            int errorCode;
            swFace1.Select4(false, null);
            swFace2.Select4(true, null);
            swAssy.AddMate3((int)swMateType_e.swMateCONCENTRIC, (int)swMateAlign_e.swMateAlignALIGNED, false, 0, 0, 0, 0, 0, 0, 0, 0, false, out errorCode);
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
        /// P1以用户设置的默认模版，新建一个零件
        /// P2在草图中创建一个圆弧slot
        /// https://www.bilibili.com/video/BV1oX4y1N74h?p=2
        /// </summary>
        public void CADCoderP2(SldWorks swApp)
        {
            //获取用户默认模版
            //get the default location of part template
            string defaultPartTemplate =
                swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            //set the solidworks document to new part document
            ModelDoc2 swModel = swApp.NewDocument(defaultPartTemplate, 0, 0, 0);
            if (swModel == null) return;
            //select front plane
            swModel.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            //insert sketch into select plane
            swModel.InsertSketch2(true);
            //create a centerpoint arc slot
            SketchSlot swSketchSlot = swModel.SketchManager.CreateSketchSlot(
                (int)swSketchSlotCreationType_e.swSketchSlotCreationType_arc,
                (int)swSketchSlotLengthType_e.swSketchSlotLengthType_CenterCenter,
                0.5, 0, 0, 0, -1, 0, 0, 1, 0, 0, -1, false);
            swModel.InsertSketch2(true);
            //de-select created centerpoint arc slot
            swModel.ClearSelection2(true);
            //zoom to fit screen in solidworks window
            swModel.ViewZoomtofit2();
        }

        /// <summary>
        /// P3在草图中创建样条曲线
        /// https://www.bilibili.com/video/BV1oX4y1N74h?p=3
        /// </summary>
        public void CADCoderP3(SldWorks swApp)
        {
            string defaultPartTemplate =
                swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            ModelDoc2 swModel = swApp.NewDocument(defaultPartTemplate, 0, 0, 0);
            if (swModel == null) return;
            swModel.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swModel.InsertSketch2(true);
            //直接使用list接收点的坐标，在使用的时候转换成数组pointArray.ToArray()
            List<double> pointArray = new List<double>();
            for (int i = 0; i < 10; i++)
            {
                //故意使用int类型损失精度，制造样条曲线的点（不在一条直线上）
                int incrementFactor = (int)(i * 0.5);
                double x = i;
                double y = x + incrementFactor;
                double z = 0;
                pointArray.Add(x);
                pointArray.Add(y);
                pointArray.Add(z);
                //create a sketch point using x,y,z variable
                swModel.SketchManager.CreatePoint(x, y, z);
            }
            swModel.ClearSelection2(true);
            //没必要了
            //Sketch swSketch = swModel.SketchManager.ActiveSketch;
            ////get all the sketch point in this active sketch and store them into our varriant type variable
            //var sketchPointArray = swSketch.GetSketchPoints2();
            //foreach (var item in sketchPointArray)
            //{
            //    SketchPoint swPoint = (SketchPoint) item;
            //}
            //根据点坐标的数组，创建样条曲线
            swModel.SketchManager.CreateSpline2(pointArray.ToArray(), true);
            swModel.InsertSketch2(true);
            swModel.ViewZoomtofit2();
        }

        /// <summary>
        /// P4在草图中创建矩形并绘制倒圆角
        /// https://www.bilibili.com/video/BV1oX4y1N74h?p=4
        /// </summary>
        public void CADCoderP4(SldWorks swApp)
        {
            string defaultPartTemplate =
                swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            ModelDoc2 swModel = swApp.NewDocument(defaultPartTemplate, 0, 0, 0);
            if (swModel == null) return;
            swModel.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swModel.InsertSketch2(true);
            //绘制矩形
            swModel.SketchManager.CreateCornerRectangle(0, 1, 0, 1, 0, 0);
            swModel.ClearSelection2(true);
            //选择点
            swModel.Extension.SelectByID2("Point1", "SKETCHPOINT", 0, 0, 0, false, 0, null, 0);
            //绘制倒圆角
            swModel.SketchManager.CreateFillet(0.1, (int)swConstrainedCornerAction_e.swConstrainedCornerDeleteGeometry);
            swModel.InsertSketch2(true);
            swModel.ViewZoomtofit2();
        }

        /// <summary>
        /// P5在草图中创建矩形并绘制倒斜角
        /// https://www.bilibili.com/video/BV1oX4y1N74h?p=5
        /// </summary>
        public void CADCoderP5(SldWorks swApp)
        {
            string defaultPartTemplate =
                swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            ModelDoc2 swModel = swApp.NewDocument(defaultPartTemplate, 0, 0, 0);
            if (swModel == null) return;
            swModel.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swModel.InsertSketch2(true);
            //绘制矩形
            swModel.SketchManager.CreateCornerRectangle(0, 1, 0, 1, 0, 0);
            swModel.ClearSelection2(true);
            //选择点
            swModel.Extension.SelectByID2("Point1", "SKETCHPOINT", 0, 0, 0, false, 0, null, 0);
            //绘制倒斜角
            swModel.SketchManager.CreateChamfer((int)swSketchChamferType_e.swSketchChamfer_DistanceDistance, 0.1, 0.2);
            swModel.InsertSketch2(true);
            //视图
            swModel.ShowNamedView2("", (int)swStandardViews_e.swFrontView);
            swModel.ViewZoomtofit2();
        }

        /// <summary>
        /// P6在草图中绘制两条直线，然后裁减/延长
        /// https://www.bilibili.com/video/BV1oX4y1N74h?p=6
        /// </summary>
        public void CADCoderP6(SldWorks swApp)
        {
            string defaultPartTemplate =
                swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            ModelDoc2 swModel = swApp.NewDocument(defaultPartTemplate, 0, 0, 0);
            if (swModel == null) return;
            swModel.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swModel.InsertSketch2(true);
            //绘制两条直线
            swModel.SketchManager.CreateLine(0, 0, 0, 1, 0, 0);
            swModel.SketchManager.CreateLine(1.5, 0, 0, 1.5, 1, 0);
            swModel.ClearSelection2(true);
            //选择两条直线
            swModel.Extension.SelectByID2("Line1", "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
            //注意参数true，表示同时添加选择
            swModel.Extension.SelectByID2("Line2", "SKETCHSEGMENT", 0, 0, 0, true, 0, null, 0);
            //裁减/延长
            swModel.SketchManager.SketchTrim((int)swSketchTrimChoice_e.swSketchTrimCorner, 0, 0, 0);
            swModel.InsertSketch2(true);
            swModel.ViewZoomtofit2();
        }

        /// <summary>
        /// P7在草图中绘制两条直线，然后偏置
        /// https://www.bilibili.com/video/BV1oX4y1N74h?p=7
        /// </summary>
        public void CADCoderP7(SldWorks swApp)
        {
            string defaultPartTemplate =
                swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            ModelDoc2 swModel = swApp.NewDocument(defaultPartTemplate, 0, 0, 0);
            if (swModel == null) return;
            swModel.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swModel.InsertSketch2(true);
            //绘制两条直线
            swModel.SketchManager.CreateLine(-0.5, 0.75, 0, -0.25, -0.5, 0);
            swModel.SketchManager.CreateLine(-0.75, -1.25, 0, 0.5, -1.25, 0);
            swModel.ClearSelection2(true);
            //选择两条直线
            swModel.Extension.SelectByID2("Line1", "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
            //注意参数true，表示同时添加选择
            swModel.Extension.SelectByID2("Line2", "SKETCHSEGMENT", 0, 0, 0, true, 0, null, 0);
            //偏置
            swModel.SketchManager.SketchOffset2(0.5, false, false, (int)swSkOffsetCapEndType_e.swSkOffsetNoCaps, (int)swSkOffsetMakeConstructionType_e.swSkOffsetDontMakeConstruction, true);
            swModel.InsertSketch2(true);
            swModel.ViewZoomtofit2();
        }

        /// <summary>
        /// P8在草图中绘制中心线和圆，然后镜像圆
        /// https://www.bilibili.com/video/BV1oX4y1N74h?p=8
        /// </summary>
        public void CADCoderP8(SldWorks swApp)
        {
            string defaultPartTemplate =
                swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            ModelDoc2 swModel = swApp.NewDocument(defaultPartTemplate, 0, 0, 0);
            if (swModel == null) return;
            swModel.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swModel.InsertSketch2(true);
            //绘制垂直的中心线和圆
            swModel.SketchManager.CreateCenterLine(0, 0, 0, 0, 1, 0);
            swModel.SketchManager.CreateCircleByRadius(-0.75, 0, 0, 0.2);
            swModel.ClearSelection2(true);
            //选择中心线和圆
            swModel.Extension.SelectByID2("Line1", "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
            //注意参数true，表示同时添加选择
            swModel.Extension.SelectByID2("Arc1", "SKETCHSEGMENT", 0, 0, 0, true, 0, null, 0);
            //镜像
            swModel.SketchMirror();
            swModel.InsertSketch2(true);
            swModel.ViewZoomtofit2();
        }

        /// <summary>
        /// P9在草图中绘制圆，然后线性阵列
        /// P10编辑线性阵列
        /// https://www.bilibili.com/video/BV1oX4y1N74h?p=9
        /// https://www.bilibili.com/video/BV1oX4y1N74h?p=10
        /// </summary>
        public void CADCoderP9P10(SldWorks swApp)
        {
            string defaultPartTemplate =
                swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            ModelDoc2 swModel = swApp.NewDocument(defaultPartTemplate, 0, 0, 0);
            if (swModel == null) return;
            swModel.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swModel.InsertSketch2(true);
            //绘制圆
            swModel.SketchManager.CreateCircleByRadius(0, 0, 0, 0.2);
            swModel.ClearSelection2(true);
            //选择圆
            swModel.Extension.SelectByID2("Arc1", "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
            //线性阵列
            swModel.SketchManager.CreateLinearSketchStepAndRepeat(3, 1, 1, 0, 0, 0, "", true, false, true, true, true);
            //P10，请单步运行观察
            //编辑线性阵列X
            swModel.SketchManager.EditLinearSketchStepAndRepeat(5, 1, 1, 0, 0, 0, "", true, false, true, true, false, "Arc1_");
            //编辑线性阵列Y
            swModel.SketchManager.EditLinearSketchStepAndRepeat(5, 4, 1, 0.75, 0, 0, "", true, false, true, true, false, "Arc1_");
            //编辑线性阵列角度
            swModel.SketchManager.EditLinearSketchStepAndRepeat(5, 4, 1, 0.75, 0.785, 1.5708, "", true, false, true, true, false, "Arc1_");
            //删除阵列数量（跳过）
            swModel.SketchManager.EditLinearSketchStepAndRepeat(5, 4, 1, 0.75, 0.785, 1.5708, "(3,2)(2,1)", true, false, true, true, false, "Arc1_");
            //显示阵列距离(Y方向)
            swModel.SketchManager.EditLinearSketchStepAndRepeat(5, 4, 1, 0.75, 0.785, 1.5708, "(3,2)(2,1)", true, true, true, true, false, "Arc1_");
            //显示与坐标轴的角度
            swModel.SketchManager.EditLinearSketchStepAndRepeat(5, 4, 1, 0.75, 0.785, 1.5708, "(3,2)(2,1)", true, true, false, true, false, "Arc1_");
            //显示整列的数量（X方向和Y方向）
            swModel.SketchManager.EditLinearSketchStepAndRepeat(5, 4, 1, 0.75, 0.785, 1.5708, "(3,2)(2,1)", true, true, false, true, true, "Arc1_");

            swModel.InsertSketch2(true);
            swModel.ViewZoomtofit2();
        }

        /// <summary>
        /// P11在草图中绘制圆，然后圆周阵列
        /// https://www.bilibili.com/video/BV1oX4y1N74h?p=11
        /// </summary>
        public void CADCoderP11(SldWorks swApp)
        {
            string defaultPartTemplate =
                swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            ModelDoc2 swModel = swApp.NewDocument(defaultPartTemplate, 0, 0, 0);
            if (swModel == null) return;
            swModel.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swModel.InsertSketch2(true);
            //绘制圆
            swModel.SketchManager.CreateCircleByRadius(0, 0, 0, 0.2);
            swModel.ClearSelection2(true);
            //选择圆
            swModel.Extension.SelectByID2("Arc1", "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
            //圆周阵列
            swModel.SketchManager.CreateCircularSketchStepAndRepeat(0.5, 0, 3, 1, true, "", true, true, true);

            swModel.InsertSketch2(true);
            swModel.ViewZoomtofit2();
        }

        /// <summary>
        /// P12打开零件
        /// https://www.bilibili.com/video/BV1oX4y1N74h?p=12
        /// </summary>
        public void CADCoderP12(SldWorks swApp)
        {
            //open a saved document
            //ModelDoc2 swModel = swApp.OpenDoc(@"E:\Videos\SolidWorks Secondary Development\SWModel\Part1.SLDPRT", (int)swDocumentTypes_e.swDocPART);
            ModelDoc2 swModel = swApp.OpenDoc6(@"E:\Videos\SolidWorks Secondary Development\SWModel\Part1.SLDPRT", (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0);
            if (swModel == null) return;
            swModel.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
        }

        /// <summary>
        /// P13在草图中绘制直线
        /// P14在草图中绘制中心线
        /// P15在草图中绘制对角矩形
        /// P15在草图中绘制中心矩形
        /// https://www.bilibili.com/video/BV1oX4y1N74h?p=13
        /// https://www.bilibili.com/video/BV1oX4y1N74h?p=14
        /// https://www.bilibili.com/video/BV1oX4y1N74h?p=15
        /// https://www.bilibili.com/video/BV1oX4y1N74h?p=16
        /// </summary>
        public void CADCoderP13141516(SldWorks swApp)
        {
            string defaultPartTemplate =
                swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            ModelDoc2 swModel = swApp.NewDocument(defaultPartTemplate, 0, 0, 0);
            if (swModel == null) return;
            swModel.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swModel.InsertSketch2(true);
            //绘制直线
            swModel.SketchManager.CreateLine(0, 0, 0, 2, 0, 0);
            //绘制中心线
            swModel.SketchManager.CreateLine(0, 0, 0, 0, 2, 0);
            //绘制对角矩形
            swModel.SketchManager.CreateCornerRectangle(0, 1, 0, 1, 0, 0);
            //绘制中心矩形
            swModel.SketchManager.CreateCenterRectangle(0, 0, 0, 1, 1, 0);
            swModel.ClearSelection2(true);
            swModel.InsertSketch2(true);
            swModel.ViewZoomtofit2();
        }

        /// <summary>
        /// P17在草图中绘制两条直线，然后延长一根直线与另一根相交
        /// https://www.bilibili.com/video/BV1oX4y1N74h?p=17
        /// </summary>
        public void CADCoderP17(SldWorks swApp)
        {
            string defaultPartTemplate =
                swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            ModelDoc2 swModel = swApp.NewDocument(defaultPartTemplate, 0, 0, 0);
            if (swModel == null) return;
            swModel.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swModel.InsertSketch2(true);
            //绘制两条直线
            swModel.SketchManager.CreateLine(-0.5, 0.75, 0, -0.25, -0.5, 0);
            swModel.SketchManager.CreateLine(-0.75, -1.25, 0, 0.5, -1.25, 0);
            swModel.ClearSelection2(true);
            //选择两条直线
            swModel.Extension.SelectByID2("Line1", "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
            //注意参数true，表示同时添加选择
            swModel.Extension.SelectByID2("Line2", "SKETCHSEGMENT", 0, 0, 0, true, 0, null, 0);
            //延长线段
            swModel.SketchManager.SketchExtend(0, 0, 0);
            swModel.InsertSketch2(true);
            swModel.ViewZoomtofit2();
        }
    }



}
