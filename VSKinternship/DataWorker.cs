using System;
using System.Data.Entity;

namespace VSKinternship
{
    class DataWorker
    {
        public bool CheckData(string key) //метод для проверки БД на совпадения ключа
        {
            using (UserContext db = new UserContext())
            {
                var users = db.Users;
                foreach (User user in users)
                {
                    if (user.Telephone == key)
                        return false;
                }
                return true;
            }
        }

        public void DataClear()
        {
            using (UserContext db = new UserContext())
            {
                DbSet<User> user = db.Users;
                foreach (User u in user)
                {
                    db.Users.Remove(u);
                }
                db.SaveChanges();
            }
        }

        public void DataOutput()
        {
            using (UserContext db = new UserContext())
            {
                Console.WriteLine("Имя Фамилия Отчество Номер телефона Адрес");
                DbSet<User> user = db.Users;
                foreach (User u in user)
                {
                    Console.WriteLine($"{u.FirstName} {u.SecondName} {u.ThirdName} {u.Telephone} {u.Address}");
                }
            }
        }
    }
}
