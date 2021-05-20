using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolConvertSqlToMongoDB.Models.eazy
{

    public class WORK_Task
    {
        public string Id { get; set; }
        public string Title { get; set; }
        public string Description { get; set; }
        public string ProjectId { get; set; }
        public string ProjectAppCode { get; set; }

        public string DateStart { get; set; }

        public string DateEnd { get; set; }

        public int SubTaskCount { get; set; }

        public int SubTaskCompleted { get; set; }

        public float SubTaskPercent { get; set; }

        public string DashboardStageId { get; set; }

        public bool IsStore { get; set; }

        public string DateCreated { get; set; }

        public string DateUpdated { get; set; }

        public bool IsCompleted { get; set; }

        public int SortOrder { get; set; }

        public bool IsConfirm { get; set; }

        public bool IsInProgress { get; set; }

        public string EmployeeId { get; set; }

        public string ContactID { get; set; }

        public string EnterpriseID { get; set; }

        public string ContractID { get; set; }

        public string CompletedDate { get; set; }
        public int CompletedPercent { get; set; }
        public int TaskPriorityID { get; set; }

        public int TaskStatusID { get; set; }

        public int TaskTypeID { get; set; }

        public bool IsReminder { get; set; }
        public int ReminderInterval { get; set; }
        public string ReminderDate { get; set; }

        public Guid OwnerID { get; set; }
        public string CreatedBy { get; set; }
        public string ModifiedBy { get; set; }
        public string CreatedDate { get; set; }
        public string ModifiedDate { get; set; }
        public string EmployeeJoinTask { get; set; }
        public string InvoiceID { get; set; }
        public bool IsPublic { get; set; }
        public bool IsService { get; set; }
        public string CampaignID { get; set; }



    }

    public class EmployeeCus
    {
        public string EmployeeID { get; set; }
        public string EmployeeCode { get; set; }
        public string Prefix { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }

        public string FullName { get; set; }

        public string Email { get; set; }

        public string Email2 { get; set; }

        public string Email3 { get; set; }

        public string Mobile { get; set; }

        public string OrganizationUnitID { get; set; }

        public string JobPositionID { get; set; }

        public string Description { get; set; }

        public bool IsSystem { get; set; }

        public int SortOrder { get; set; }

        public string Password { get; set; }

        public bool IsPublic { get; set; }

        public string HomeAddressID { get; set; }

        public DateTime BirthDate { get; set; }

        public string HomeAddress { get; set; }

        public bool Inactive { get; set; }

        public bool IsLeave { get; set; }

        public DateTime LeaveFromDate { get; set; }
        public string Avatar { get; set; }
        public string Facebook { get; set; }

        public string YM { get; set; }

        public string Skype { get; set; }

        public string Zalo { get; set; }
        public bool IsWorking { get; set; }
        public string CardNo { get; set; }
        public string UserID2 { get; set; }
        public string CardNo2 { get; set; }
        public Decimal BasicPay { get; set; }
        public Decimal Allowance1 { get; set; }
        public Decimal Allowance2 { get; set; }
        public Decimal Insurance { get; set; }
        public string BankAccount { get; set; }
        public string BankBranchName { get; set; }
        public string TelExt { get; set; }
        public bool IsTemp { get; set; }
        public bool BasicPayByDay { get; set; }
        public string IDCard { get; set; }
        public DateTime IDCardDate { get; set; }
        public string IDCardBy { get; set; }
        public bool IsNotAT { get; set; }
        public string AccountGroupId { get; set; }
        public string AccountStatus { get; set; }
        public bool Active { get; set; }
        public string CodeReset { get; set; }
        public DateTime DateUpdated { get; set; }
        public string Fax { get; set; }
        public bool IsDeleted { get; set; }
        public bool IsDisplayInHR { get; set; }
        public string Name { get; set; }
        public string PersonEmail1 { get; set; }
        public string Phone { get; set; }
        public string dutyid { get; set; }
        public bool isAdmin { get; set; }
        public string locationid { get; set; }
        public string organizationid { get; set; }
    }
    public class AccountRole
    {
        public string Id { get; set; }
        public string EmployeeId { get; set; }
        public string RoleId { get; set; }


    }
    public class First_Last_Name
    {
        public List<string> LstFirstName { get; set; } = new List<string>();
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string FullName { get; set; }

    }
}