//#region License
///*
//Copyright © 2014-2018 European Support Limited

//Licensed under the Apache License, Version 2.0 (the "License")
//you may not use this file except in compliance with the License.
//You may obtain a copy of the License at 

//http://www.apache.org/licenses/LICENSE-2.0 

//Unless required by applicable law or agreed to in writing, software
//distributed under the License is distributed on an "AS IS" BASIS, 
//WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. 
//See the License for the specific language governing permissions and 
//limitations under the License. 
//*/
//#endregion

//using Amdocs.Ginger.Plugin.Core;
//using GingerTestHelper;
//using Microsoft.VisualStudio.TestTools.UnitTesting;
//using StandAloneActions;
//using System;

//namespace StandAloneActionsTest
//{
//    [TestClass]
//    public class LargeExcelTest
//    {
//        string EXCEL_FILE_NAME = TestResources.GetTestResourcesFile("LargeExcelDataTest.xlsx");
           
//        [TestMethod]
//        public void ReadExcelCellRow4ColumnD()
//        {
//            //Arrange
//            ExcelAction x = new ExcelAction();
//            GingerAction GA = new GingerAction("Excel");

//            //Act
//            x.ReadExcelCell(ref GA, EXCEL_FILE_NAME, "Sheet1", "#4", "#D");

//            //Assert
//            Assert.AreEqual("Darrin", GA.Output.Values[0].ValueString, "Row 4 Col D = Darrin");
//        }

//        [TestMethod]
//        public void ReadExcelCellMatchingFirstName()
//        {
//            //Arrange
//            ExcelAction x = new ExcelAction();
//            GingerAction GA = new GingerAction("Excel");

//            //Act
//            x.ReadExcelCell(ref GA, EXCEL_FILE_NAME, "Sheet1", "FirstName='Brosina'", "RowID");
            
//            //Assert            
//            Assert.AreEqual("6", GA.Output.Values[0].ValueString, "First=Brosina RowID=6");
//        }

//        [TestMethod]
//        public void ReadExcelRow17()
//        {
//            //Arrange
//            ExcelAction x = new ExcelAction();
//            GingerAction GA = new GingerAction("Excel");

//            //Act
//            x.ReadExcelRow(ref GA, EXCEL_FILE_NAME, "Sheet1", "#17", string.Empty);     // if no columns are specified read them all       
            
//            //Assert
//            Assert.AreEqual("16", GA.Output.Values[0].ValueString, "Row 17 RowID=16");
//            Assert.AreEqual("US-2015-118983", GA.Output.Values[1].ValueString, "Row 17 OrderID=US-2015-118983");
//            Assert.AreEqual("HP-14815", GA.Output.Values[2].ValueString, "Row 17 CustomerID=HP-14815");
//            Assert.AreEqual("Harold", GA.Output.Values[3].ValueString, "Row 17 FirstName=Harold");
//            Assert.AreEqual("Pawlan", GA.Output.Values[4].ValueString, "Row 17 LastName=Pawlan");
//        }
        
//        [TestMethod]
//        public void ReadExcelRow111Get3Columns()
//        {
//            //Arrange
//            ExcelAction x = new ExcelAction();
//            GingerAction GA = new GingerAction("Excel");

//            //Act
//            x.ReadExcelRow(ref GA, EXCEL_FILE_NAME, "Sheet1", "#111", "#D, #5, ProductName");  // Read row 3 columns #D,#6, ProductName
            
//            //Assert
//            Assert.AreEqual("Pete", GA.Output.Values[0].ValueString, "Row 111 FirstName=Pete");
//            Assert.AreEqual("Armstrong", GA.Output.Values[1].ValueString, "Row 111 LastName=Armstrong");
//            Assert.AreEqual("Logitech Gaming G510s - Keyboard", GA.Output.Values[2].ValueString, "Row 111 ProductName=Logitech Gaming G510s - Keyboard");
//            Assert.AreEqual(3, GA.Output.Values.Count);
//        }

//        [TestMethod]
//        public void ReadExcelCellAsProductName()
//        {
//            //Arrange
//            ExcelAction x = new ExcelAction();
//            GingerAction GA = new GingerAction("Excel");

//            //Act
//            x.ReadExcelRow(ref GA, EXCEL_FILE_NAME, "Sheet1", "ProductName='Newell 322'", "#D,#5, ProductName");  // Read first row where Used = No and get the ID

//            //Assert
//            Assert.AreEqual("Brosina", GA.Output.Values[0].ValueString, "Row 7 ID=Brosina");
//            Assert.AreEqual("Hoffman", GA.Output.Values[1].ValueString, "Row 7 First=Hoffman");
//            Assert.AreEqual("Newell 322", GA.Output.Values[2].ValueString, "Row 7 ProductName=Newell 322");
//        }

//        [TestMethod]
//        public void ReadExcelRowWithComplexCondition()
//        {
//            //Arrange
//            ExcelAction x = new ExcelAction();
//            GingerAction GA = new GingerAction("Excel");

//            //Act
//            x.ReadExcelRow(ref GA, EXCEL_FILE_NAME, "Sheet1", "Sales>'1700' and Quantity='9'", string.Empty);   // Read the first row when ID>30 and Used=No, get column: A,B,D,E

//            //Assert
//            Assert.AreEqual("11", GA.Output.Values[0].ValueString, "Row 12 RowID=11");
//            Assert.AreEqual("CA-2014-115812", GA.Output.Values[1].ValueString, "Row 12 OrderID=CA-2014-115812");
//            Assert.AreEqual("BH-11710", GA.Output.Values[2].ValueString, "Row 12 CustomerID=BH-11710");
//            Assert.AreEqual("Brosina", GA.Output.Values[3].ValueString, "Row 12 FirstName=Brosina");
//            Assert.AreEqual("Hoffman", GA.Output.Values[4].ValueString, "Row 12 LastName=Hoffman");
//        }

//        [TestMethod]
//        public void ReadUpdatePassExcelCellWithRowNumber()
//        {
//            //Arrange
//            ExcelAction x = new ExcelAction();
//            GingerAction GA = new GingerAction("Excel");

//            //Act
//            x.ReadExcelAndUpdate(ref GA, EXCEL_FILE_NAME, "Sheet1", "#13", string.Empty, "CustomerID='BH-11711', #E='Maddox', ProductID='TEC-PH-10002034'");      

//            //Assert
//            Assert.AreEqual(Convert.ToString(true), GA.Output.Values[0].ValueString);
//        }

//        [TestMethod]
//        public void ReadUpdateFailExcelCellWithRowNumber()
//        {
//            //Arrange
//            ExcelAction x = new ExcelAction();
//            GingerAction GA = new GingerAction("Excel");

//            //Act
//            x.ReadExcelAndUpdate(ref GA, EXCEL_FILE_NAME, "Sheet1", "#2070", string.Empty, "RowID='2070', #I='North Carolina', ShipMode='Third Class'");

//            //Assert
//            Assert.IsFalse(string.IsNullOrEmpty(GA.Errors));            
//        }

//        [TestMethod]
//        public void ReadUpdatePassExcelCellWithRowNumberAndGetColumnsData()
//        {
//            //Arrange
//            ExcelAction x = new ExcelAction();
//            GingerAction GA = new GingerAction("Excel");

//            //Act
//            x.ReadExcelAndUpdate(ref GA, EXCEL_FILE_NAME, "Sheet1", "#33", "#E, #13, PostalCode", "#O='10', SubCategory='Art'");

//            //Assert
//            Assert.AreEqual("Blumstein", GA.Output.Values[0].ValueString, "Row 32 LastName=Blumstein");
//            Assert.AreEqual("BOSTON Model 1800 Electric Pencil Sharpeners, Putty/Woodgrain", GA.Output.Values[1].ValueString, "Row 32 ProductName=BOSTON Model 1800 Electric Pencil Sharpeners, Putty/Woodgrain");
//            Assert.AreEqual("19140", GA.Output.Values[2].ValueString, "Row 32 OrderID=19140");            
//        }

//        [TestMethod]
//        public void AppendNewRowSimple()
//        {
//            //Arrange
//            ExcelAction x = new ExcelAction();
//            GingerAction GA = new GingerAction("Excel");

//            x.AppendData(ref GA, EXCEL_FILE_NAME, "Sheet1", "'2060', 'CA-2015-4444444', 'CW-120000', 'Carol', 'Norling', 'Home Office'");
//            int index = Convert.ToInt32(GA.Output.Values[0].ValueString);
            
//            //Assert
//            Assert.IsTrue(index > 0);
//        }

//        [TestMethod]
//        public void AppendNewRowByColumnHeading()
//        {
//            //Arrange
//            ExcelAction x = new ExcelAction();
//            GingerAction GA = new GingerAction("Excel");

//            x.AppendData(ref GA, EXCEL_FILE_NAME, "Sheet1", "FirstName='Alejandro', LastName='Burns', ProductID='FUR-FU-10004848', SubCategory='Art'");
//            int index = Convert.ToInt32(GA.Output.Values[0].ValueString);

//            //Assert
//            Assert.IsTrue(index > 0);
//        }

//        [TestMethod]
//        public void AppendNewRowByColumnName()
//        {
//            //Arrange
//            ExcelAction x = new ExcelAction();
//            GingerAction GA = new GingerAction("Excel");

//            x.AppendData(ref GA, EXCEL_FILE_NAME, "Sheet1", "#D='Zuschuss', #City='Orem'");
//            int index = Convert.ToInt32(GA.Output.Values[0].ValueString);

//            //Assert
//            Assert.IsTrue(index > 0);
//        }

//        [TestMethod]
//        public void WriteExcel()
//        {
//            //Arrange
//            ExcelAction x = new ExcelAction();
//            GingerAction GA = new GingerAction("Excel");

//            //Act
//            x.WriteExcel(ref GA, EXCEL_FILE_NAME, "Sheet1", 2056, "E", "Write Cell");
            
//            //Assert
//            Assert.IsTrue(string.IsNullOrEmpty(GA.Errors));
//        }       
//    }
//}
