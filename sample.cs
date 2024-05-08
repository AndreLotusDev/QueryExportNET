using Dapper;
using Npgsql;
using OfficeOpenXml;

var dbOne = "CONNECTION";
var dbTwo = "CONNECTION";

var query = @"select u.""Id"", u.""FirstName"", u.""LastName"", u.""Email"", uc.""ClaimValue"" from ""User"" u left join ""UserClaims"" uc on u.""Id"" = uc.""UserId"";";

using var connection = new NpgsqlConnection(dbOne);

var users = connection.Query<QueryUser>(query).ToList();

var querySalesThatAUserCanView = @"select uspa.user_id as user_id, ud.user_id as can_view from user_sales_person_access uspa left join user_department ud on uspa.sales_person_id = ud.id";

using var connection2 = new NpgsqlConnection(dbTwo);

var salesThatAUserCanView = connection2.Query<QuerySales>(querySalesThatAUserCanView).ToList();

var groupingByUserId = users.GroupBy(x => x.Id)
    .ToDictionary(x => x.Key, x => x.Select(s => new UserInfo()
    {
        UserId = s.Id,
        FullName = $"{s.FirstName} {s.LastName}",
        Claims = x.Select(s => s.ClaimValue).Distinct().ToList(),
        Sales = salesThatAUserCanView.Where(y => y.user_id == s.Id).Select(y => new UserInfo()
        {
            UserId = y.can_view,
            FullName = users.FirstOrDefault(z => z.Id == y.can_view)?.FirstName + " " + users.FirstOrDefault(z => z.Id == y.can_view)?.LastName
        }).ToList()
    }).ToList());

Console.WriteLine("Everything finished");

var excel = new ExcelPackage();
var ws = excel.Workbook.Worksheets.Add("Sheet1");

ws.Cells["A1"].Value = "User Id";
ws.Cells["B1"].Value = "Full Name";
ws.Cells["C1"].Value = "Claims";
ws.Cells["D1"].Value = "Sales Person";

var startRowPosition = 2;
//Add column user id, full name, and the column with the name of sales persons
foreach (var (key, value) in groupingByUserId)
{
    ws.Cells["A" + startRowPosition].Value = key;
    ws.Cells["B" + startRowPosition].Value = value.FirstOrDefault()?.FullName;
    ws.Cells["C" + startRowPosition].Value = string.Join(", ", value.FirstOrDefault()?.Claims);
    ws.Cells["D" + startRowPosition].Value = string.Join(", ", value.FirstOrDefault()?.Sales.Select(x => x.FullName).ToList());
    startRowPosition++;
}

excel.SaveAs(new FileInfo("C:\\Users\\andrs\\Desktop\\ExportSalesAndUserClaims.xlsx"));

class UserInfo
{
    public string UserId { get; set; }
    public string FullName { get; set; }

    public List<string> Claims { get; set; }
    public List<UserInfo> Sales { get; set; } = new();
}

class QueryUser
{
    public string FirstName { get; set; }
    public string LastName { get; set; }
    public string Email { get; set; }
    public string ClaimValue { get; set; }
    public string Id { get; set; }
}

class QuerySales 
{
    public string user_id { get; set; }
    public string can_view { get; set; }
}
