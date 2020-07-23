using System.Collections.Generic;

namespace DrinksCaculator.DrinksObject.Member
{
    public class Member : IMemberBase
    {
        public string No { get; set; }

        public string StaffName { get; set; }

        public string[] NickName { get; set; }
    }

    public class ApexMembers
    {
        public IEnumerable<Member> Staffs { get; set; }
    }
}
