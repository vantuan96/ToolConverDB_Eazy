using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolConvertSqlToMongoDB.Models
{
    public class tblCardGroup
    {
        public string Id { get; set; }
        public string CardGroupID { get; set; }

        public string CardGroupCode { get; set; }

        public string CardGroupName { get; set; }

        public string Description { get; set; }

        public int CardType { get; set; }

        public string VehicleGroupID { get; set; }

        public string LaneIDs { get; set; }

        public string DayTimeFrom { get; set; }

        public string DayTimeTo { get; set; }

        public int Formulation { get; set; }

        public int EachFee { get; set; }

        public int Block0 { get; set; }

        public int Time0 { get; set; }

        public int Block1 { get; set; }

        public int Time1 { get; set; }

        public int Block2 { get; set; }

        public int Time2 { get; set; }

        public int Block3 { get; set; }

        public int Time3 { get; set; }

        public int Block4 { get; set; }

        public int Time4 { get; set; }

        public int Block5 { get; set; }

        public int Time5 { get; set; }

        public string TimePeriods { get; set; }

        public string Costs { get; set; }

        public bool Inactive { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int SortOrder { get; set; }

        public bool IsHaveMoneyExcessTime { get; set; }

        public bool EnableFree { get; set; }

        public int FreeTime { get; set; }

        public bool IsCheckPlate { get; set; }

        public bool IsHaveMoneyExpiredDate { get; set; }
    }
    public class tblCardGroup1
    {
        [Key]
        public System.Guid CardGroupID { get; set; }

        public string CardGroupCode { get; set; }

        public string CardGroupName { get; set; }

        public string Description { get; set; }

        public int CardType { get; set; }

        public string VehicleGroupID { get; set; }

        public string LaneIDs { get; set; }

        public string DayTimeFrom { get; set; }

        public string DayTimeTo { get; set; }

        public int Formulation { get; set; }

        public int EachFee { get; set; }

        public int Block0 { get; set; }

        public int Time0 { get; set; }

        public int Block1 { get; set; }

        public int Time1 { get; set; }

        public int Block2 { get; set; }

        public int Time2 { get; set; }

        public int Block3 { get; set; }

        public int Time3 { get; set; }

        public int Block4 { get; set; }

        public int Time4 { get; set; }

        public int Block5 { get; set; }

        public int Time5 { get; set; }

        public string TimePeriods { get; set; }

        public string Costs { get; set; }

        public bool Inactive { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int SortOrder { get; set; }

        public bool IsHaveMoneyExcessTime { get; set; }

        public bool EnableFree { get; set; }

        public int FreeTime { get; set; }

        public bool IsCheckPlate { get; set; }

        public bool IsHaveMoneyExpiredDate { get; set; }
    }
}
