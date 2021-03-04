using System;
using System.ComponentModel.DataAnnotations;
using System.Data.Entity;

namespace VSKinternship
{
    public class User
    {
        public string FirstName { get; set; }
        public string SecondName { get; set; }
        public string ThirdName { get; set; }
        [Key]
        public string Telephone { get; set; }
        public string Address { get; set; }
    }

    public class UserContext : DbContext
    {
        public UserContext() :
            base("UserDB")
        { }

        public DbSet<User> Users { get; set; }
    }
}
