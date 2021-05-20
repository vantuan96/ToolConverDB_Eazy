using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolConvertSqlToMongoDB.Models
{
    [BsonIgnoreExtraElements]
    [Table("tblCustomer")]
    public class tblCustomer
    {
        public string Id { get; set; }
        public string CustomerID { get; set; }

        public string CustomerCode { get; set; }

        public string CustomerName { get; set; }

        public string Address { get; set; }

        public string IDNumber { get; set; }

        public string Mobile { get; set; }

        public string CustomerGroupID { get; set; }

        public string Description { get; set; }

        public bool EnableAccount { get; set; }

        public string Account { get; set; }

        public string Password { get; set; }

        public string Avatar { get; set; }

        public bool Inactive { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int SortOrder { get; set; }

        public string CompartmentId { get; set; }

        public string AccessLevelID { get; set; }

        public string Finger1 { get; set; }

        public string Finger2 { get; set; }

        public int UserIDofFinger { get; set; }

        public System.DateTime AccessExpireDate { get; set; }

        public string DevPass { get; set; }
        public Nullable<System.DateTime> ContractStartDate { get; set; }
        public Nullable<System.DateTime> ContractEndDate { get; set; }

        public string AddressUnsign { get; set; }

        public string MS_PersonGroupId { get; set; }
        public int UserFaceId { get; set; }

        //Update aeon hp 19/11/2020
        public string Plate1 { get; set; }

        public string PlateUnsign1 { get; set; }

        public string VehicleName1 { get; set; }

        public string Plate2 { get; set; }
        public string PlateUnsign2 { get; set; }

        public string VehicleName2 { get; set; }

        public string Plate3 { get; set; }

        public string PlateUnsign3 { get; set; }

        public string VehicleName3 { get; set; }
    }

}
