using Bogus;

public class Person
{
    public string FirstName { get; set; } = string.Empty;
    public string LastName { get; set; } = string.Empty;
    public int Age { get; set; }
    public string Email { get; set; } = string.Empty;
    public string PhoneNumber { get; set; } = string.Empty;
    public string Address { get; set; } = string.Empty;

    private static readonly Faker<Person> faker = new Faker<Person>()
        .RuleFor(p => p.FirstName, f => f.Name.FirstName())
        .RuleFor(p => p.LastName, f => f.Name.LastName())
        .RuleFor(p => p.Age, f => f.Random.Int(18, 100))
        .RuleFor(p => p.Email, f => f.Internet.Email())
        .RuleFor(p => p.PhoneNumber, f => f.Phone.PhoneNumber())
        .RuleFor(p => p.Address, f => f.Address.StreetAddress());

    /// <summary>
    /// Generates a random person with realistic data.
    /// </summary>
    /// <returns>A randomly generated <see cref="Person"/> instance.</returns>
    public static Person Random() => faker.Generate();
}