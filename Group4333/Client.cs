using System;
using System.ComponentModel.DataAnnotations;

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