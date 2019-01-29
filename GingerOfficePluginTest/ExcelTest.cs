#region License
/*
Copyright © 2014-2018 European Support Limited

Licensed under the Apache License, Version 2.0 (the "License")
you may not use this file except in compliance with the License.
You may obtain a copy of the License at 

http://www.apache.org/licenses/LICENSE-2.0 

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS, 
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. 
See the License for the specific language governing permissions and 
limitations under the License. 
*/
#endregion

using Amdocs.Ginger.Plugin.Core;
using GingerTestHelper;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using StandAloneActions;
using System;

namespace StandAloneActionsTest
{
    [TestClass]
    public class ExcelTest
    {
        string EXCEL_FILE_NAME = TestResources.GetTestResourcesFile("test1.xlsx");
           
        [TestMethod]
        public void ReadExcelCellRow3ColumnB()
        {
            //Arrange
            ExcelService x = new ExcelService();
            GingerAction GA = new GingerAction();

            //Act
            x.ReadExcelCell(GA, EXCEL_FILE_NAME, "Sheet1", "#3", "#B");

            //Assert
            //Assert.AreEqual("Moshe", GA.Output.Values[0].ValueString, "Row 3 Col B = Moshe");
        }

        [TestMethod]
        public void ReadExcelCellRow3ColumnColumnCount()
        {
            //Arrange
            ExcelService x = new ExcelService();
            GingerAction GA = new GingerAction();

            //Act
            x.ReadExcelCell(GA, EXCEL_FILE_NAME, "Sheet1", "#3", "#B");

            //Assert
            //Assert.AreEqual(1, GA.Output.Values.Count);
        }

        [TestMethod]
        public void ReadExcelCellWhereFirstEquelMosheGetID()
        {
            //Arrange
            ExcelService x = new ExcelService();
            GingerAction GA = new GingerAction();

            //Act
            x.ReadExcelCell(GA, EXCEL_FILE_NAME, "Sheet1", "First='Moshe'", "ID");
            
            //Assert            
            // Assert.AreEqual("24", GA.Output.Values[0].ValueString, "First=Moshe ID=24");
        }

        [TestMethod]
        public void ReadExcelRow2()
        {
            //Arrange
            ExcelService x = new ExcelService();
            GingerAction GA = new GingerAction();

            //Act
            x.ReadExcelRow(GA, EXCEL_FILE_NAME, "Sheet1", "#2", string.Empty);     // if no columns are specified read them all       
            
            //Assert
            //Assert.AreEqual("12", GA.Output.Values[0].ValueString, "Row 2 ID=12");
            //Assert.AreEqual("David", GA.Output.Values[1].ValueString, "Row 2 First=David");
            //Assert.AreEqual("Cohen", GA.Output.Values[2].ValueString, "Row 2 Last=Cohen");
            //Assert.AreEqual("923646", GA.Output.Values[3].ValueString, "Row 2 Phone=923646");
            //Assert.AreEqual("Yes", GA.Output.Values[4].ValueString, "Row 2 Used=Yes");
        }
        
        [TestMethod]
        public void ReadExcelRow3GetColumns_1_2_4()
        {
            //Arrange
            ExcelService x = new ExcelService();
            GingerAction GA = new GingerAction();

            //Act
            x.ReadExcelRow(GA, EXCEL_FILE_NAME, "Sheet1", "#3","#1,#2,#4");  // Read row 3 columns 1,2,4
            
            //Assert            
            //Assert.AreEqual("24", GA.Output.Values[0].ValueString, "Row 3 ID=24");
            //Assert.AreEqual("Moshe", GA.Output.Values[1].ValueString, "Row 3 First=Moshe");
            //Assert.AreEqual("73769", GA.Output.Values[2].ValueString, "Row 3 Phone=73769");            
        }

        [TestMethod]
        public void ReadExcelRow3GetColumnsCount()
        {
            //Arrange
            ExcelService x = new ExcelService();
            GingerAction GA = new GingerAction();

            //Act
            x.ReadExcelRow(GA, EXCEL_FILE_NAME, "Sheet1", "#3", "#1,#2,#4");  // Read row 3 columns 1,2,4
                                                                        //Assert
            // Assert.AreEqual(3, GA.Output.Values.Count);
        }

        [TestMethod]
        public void ReadExcelCellWhereUsedIsNoAndGetID()
        {
            //Arrange
            ExcelService x = new ExcelService();
            GingerAction GA = new GingerAction();

            //Act
            x.ReadExcelRow(GA, EXCEL_FILE_NAME, "Sheet1", "Used='No'", "ID");  // Read first row where Used = No and get the ID

            //Assert
            // Assert.AreEqual("24", GA.Output.Values[0].ValueString, "Used=No is ID=24");
        }

        [TestMethod]
        public void ReadExcelRowWithConditionIDMoreThan30AndUsedIsNo()
        {
            //Arrange
            ExcelService x = new ExcelService();
            GingerAction GA = new GingerAction();

            //Act
            x.ReadExcelRow(GA, EXCEL_FILE_NAME, "Sheet1", "ID>'30' and Used='No'", "#A, #2, Phone, #5");   // Read the first row when ID>30 and Used=No, get column: A,B,D,E

            //Assert
            //Assert.AreEqual("32", GA.Output.Values[0].ValueString, "Row 4 ID=32");
            //Assert.AreEqual("Dana", GA.Output.Values[1].ValueString, "Row 4 First=Dana");
            //Assert.AreEqual("878375", GA.Output.Values[2].ValueString, "Row 4 Phone=878375");
            //Assert.AreEqual("No", GA.Output.Values[3].ValueString, "Row 4 Used=No");
        }

        [TestMethod]
        public void ReadUpdatePassExcelCellWithRowNoGetColumnName()
        {
            //Arrange
            ExcelService x = new ExcelService();
            GingerAction GA = new GingerAction();

            //Act
            x.ReadExcelAndUpdate(GA, EXCEL_FILE_NAME, "Sheet1", "#3", "#1,#2,#4", "ID='12', #C='John', First='Dave'");      //1   John is optional

            //Assert    
            //Assert.AreEqual("24", GA.Output.Values[0].ValueString, "Row 3 ID=24");
            //Assert.AreEqual("Moshe", GA.Output.Values[1].ValueString, "Row 3 First=Moshe");
            //Assert.AreEqual("73769", GA.Output.Values[2].ValueString, "Row 3 Phone=73769");
        }

        [TestMethod]
        public void ReadUpdateFailExcelCellWithRowNo()
        {
            ///Arrange
            ExcelService x = new ExcelService();
            GingerAction GA = new GingerAction();

            //Act
            x.ReadExcelAndUpdate(GA, EXCEL_FILE_NAME, "Sheet1", "#5", string.Empty, "ID=5, #C='Thomas', First='Sarah'");

            //Assert
            Assert.IsFalse(string.IsNullOrEmpty(GA.Errors));
        }
        
        [TestMethod]
        public void AppendNewRowSimple()
        {
            //Arrange
            ExcelService x = new ExcelService();
            GingerAction GA = new GingerAction();
            
            x.AppendData(GA, EXCEL_FILE_NAME, "Sheet1", "'55', 'John', 'Smith'");
            // int index = Convert.ToInt32(GA.Output.Values[0].ValueString);
            
            //Assert
            // Assert.IsTrue(index > 0);
        }

        [TestMethod]
        public void AppendNewRowByColumnHeading()
        {
            //Arrange
            ExcelService x = new ExcelService();
            GingerAction GA = new GingerAction();

            x.AppendData(GA, EXCEL_FILE_NAME, "Sheet1", "Last='aaa', Used='No'");
            // int index = Convert.ToInt32(GA.Output.Values[0].ValueString);

            //Assert
            // Assert.IsTrue(index > 0);
        }

        [TestMethod]
        public void AppendNewRowByColumnName()
        {
            //Arrange
            ExcelService x = new ExcelService();
            GingerAction GA = new GingerAction();

            x.AppendData(GA, EXCEL_FILE_NAME, "Sheet1", "#B='ColName', #E='Yes'");
            // int index = Convert.ToInt32(GA.Output.Values[0].ValueString);

            //Assert
            // Assert.IsTrue(index > 0);
        }

        [TestMethod]
        public void WriteExcelPass()
        {
            //Arrange
            ExcelService x = new ExcelService();
            GingerAction GA = new GingerAction();

            //Act
            x.WriteExcel(GA, EXCEL_FILE_NAME, "Sheet1", 7, "B", "Write Cell");

            //Assert
            Assert.IsTrue(string.IsNullOrEmpty(GA.Errors));
        }
    }
}
