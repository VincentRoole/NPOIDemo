using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using Maticsoft.DBUtility;

namespace NPOIDemo
{
    public partial class DemoTest : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }
         
        protected void btn_Export_Click(object sender, EventArgs e)
        {
            DataTable dt = CreateDataTable();
            IList<NPOIModel> list = new List<NPOIModel>();
            list.Add(new NPOIModel(dt, "name;sex;workyear;position;hiredate;skill;hobby;avg", "测试1",
                "姓名#性别#工作情况 工作年限,职位#入职日期#其他信息 能力 技能,爱好#其他信息 年龄", "测试1"));
            dt = CreateDataTable();
            list.Add(new NPOIModel(dt, null, "测试2", null, "测试2"));

            dt = FarmerList();
            list.Add(new NPOIModel(dt, "rowid;tname;vname;Name;IdNumber;education;political;mz;Household;HouseholdType;Phone;Domicile", "测试3",
                "所属乡镇 序号,乡镇,行政村#所在村 基本信息 姓名,身份证号,文化程度,政治面貌,民族,户籍,户籍性质,手机,现居住地", "崇明县农业从业人员信息统计表"));
          
            NPOIHelper.Export(null, list); 
        }

        private DataTable FarmerList() {
            string sql = @"SELECT TOP 20 ROW_NUMBER()OVER(ORDER BY a.Id) rowid,b.Name tname,c.Name vname,a.Name,a.IdNumber,d.Name education,e.Name political,
                        '汉族' mz,a.Household,a.HouseholdType,a.Phone,Domicile FROM dbo.tb_FarmerInfo a
                        LEFT JOIN dbo.View_Area b ON b.id=a.CurrentArea
                        LEFT JOIN dbo.View_Area c ON c.id=a.InCome
                        LEFT JOIN dbo.tb_BaseData d ON d.Id=a.Education
                        LEFT JOIN dbo.tb_BaseData e ON e.Id=a.Political
                        WHERE a.InCome IS NOT NULL";

            return DbHelperSQL.Query(sql).Tables[0];
        }

        private DataTable CreateDataTable()
        {
            DataTable dt = new DataTable();
            DataColumn col = new DataColumn("name", typeof(string));
            dt.Columns.Add(col);                   
            col = new DataColumn("sex", typeof(string));
            dt.Columns.Add(col);
            col = new DataColumn("avg", typeof(int));
            dt.Columns.Add(col);
            col = new DataColumn("mobilephone", typeof(decimal));
            dt.Columns.Add(col);
            col = new DataColumn("workyear", typeof(int));
            dt.Columns.Add(col);
            col = new DataColumn("position", typeof(string));
            dt.Columns.Add(col);
            col = new DataColumn("skill", typeof(string));
            dt.Columns.Add(col);
            col = new DataColumn("hobby", typeof(string));
            dt.Columns.Add(col);
            col = new DataColumn("hiredate", typeof(DateTime));
            dt.Columns.Add(col);

            DataRow rw = dt.NewRow();
            rw["name"] = "Karl";
            rw["sex"] = "男";
            rw["avg"] = 20;
            rw["mobilephone"] = 15890234430;
            rw["workyear"] = 3;
            rw["position"] = "程序员";
            rw["skill"] = ".net,oracel,sqlserver,html5,css3";
            rw["hobby"] = "打篮球，运动";
            rw["hiredate"] = DateTime.Now;
            dt.Rows.Add(rw);
            rw = dt.NewRow();
            rw["name"] = "rose";
            rw["sex"] = "女";
            rw["avg"] = 23;
            rw["mobilephone"] = 17890234430;
            rw["workyear"] = 4;
            rw["position"] = "美工";
            rw["skill"] = "ps";
            rw["hobby"] = "唱歌，跳舞，做饭";
            rw["hiredate"] = DateTime.Now;
            dt.Rows.Add(rw);
            rw = dt.NewRow();
            rw["name"] = "boss";
            rw["sex"] = "男";
            rw["avg"] = 45;
            rw["mobilephone"] = 13890239980;
            rw["workyear"] = 20;
            rw["position"] = "老总";
            rw["skill"] = "人际关系，企业管理";
            rw["hobby"] = "爬山，摄影，跑步，游泳";
            rw["hiredate"] = DateTime.Now;
            dt.Rows.Add(rw); 
            return dt;
        }
    }
}