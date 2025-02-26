using LazyData.Excel;

List<Person> people = [];
for (int i = 0; i < 100; i++)
{
    people.Add(Person.Random());
}

people.SaveAsExcelFile("people.xlsx");

var peopleFromFile = new List<Person>().FromExcelFilePath("people.xlsx");

foreach (var person in peopleFromFile)
{
    Console.WriteLine($"{person.FirstName} {person.LastName} ({person.Age})");
}