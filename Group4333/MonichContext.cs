using System;
using System.Data.Entity; // Если используешь .NET Framework

namespace Group4333
{
    // Класс контекста — это "пульт управления" твоей базой данных
    public class MonichContext : DbContext
    {
        // Конструктор: "MonichConnection" — это имя строки из App.config (создадим на след. шаге)
        public MonichContext() : base("name=MonichConnection")
        {
        }

        public DbSet<Clients> Clients { get; set; }
    }

    // Класс сущности (таблица Clients из твоей БД)
    public class Clients
    {
        public int ID { get; set; }
        public string ClientCode { get; set; }
        public string FullName { get; set; }
        public DateTime BirthDate { get; set; }
        public string IndexCode { get; set; }
        public string City { get; set; }
        public string Street { get; set; }
        public string House { get; set; }
        public string Apartment { get; set; }
        public string Email { get; set; }
        public int Age { get; set; }
    }
}