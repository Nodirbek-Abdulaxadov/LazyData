namespace LazyData;

internal static class PluralizationHelper
{
    internal static string ToPlural(string className)
    {
        if (string.IsNullOrEmpty(className))
            return className;

        // Handle common irregular nouns
        var irregularPlurals = new Dictionary<string, string>
        {
            {"Person", "People"},
            {"Child", "Children"},
            {"Man", "Men"},
            {"Woman", "Women"},
            {"Mouse", "Mice"},
            {"Goose", "Geese"},
            {"Foot", "Feet"},
            {"Tooth", "Teeth"},
            {"Cactus", "Cacti"},
            {"Focus", "Foci"},
            {"Fungus", "Fungi"},
            {"Nucleus", "Nuclei"},
            {"Syllabus", "Syllabi"},
            {"Analysis", "Analyses"},
            {"Diagnosis", "Diagnoses"},
            {"Oasis", "Oases"},
            {"Thesis", "Theses"},
            {"Crisis", "Crises"},
            {"Phenomenon", "Phenomena"},
            {"Criterion", "Criteria"},
            {"Datum", "Data"},
            {"Alumnus", "Alumni"},
            {"Appendix", "Appendices"},
            {"Index", "Indices"},
            {"Matrix", "Matrices"},
            {"Ox", "Oxen"},
            {"Vortex", "Vortices"},
            {"Elf", "Elves"},
            {"Calf", "Calves"},
            {"Knife", "Knives"},
            {"Leaf", "Leaves"},
            {"Life", "Lives"},
            {"Wife", "Wives"},
            {"Wolf", "Wolves"},
            {"Shelf", "Shelves"},
            {"Self", "Selves"},
            {"Loaf", "Loaves"},
            {"Scarf", "Scarves"},
            {"Thief", "Thieves"},
            {"Half", "Halves"},
            {"Tomato", "Tomatoes"},
            {"Potato", "Potatoes"},
            {"Hero", "Heroes"},
            {"Echo", "Echoes"},
            {"Torpedo", "Torpedoes"},
            {"Embryo", "Embryos"},
            {"Embargo", "Embargoes"}
        };

        if (irregularPlurals.ContainsKey(className))
        {
            return irregularPlurals[className];
        }

        // Handle general rules
        if (className.EndsWith("y", StringComparison.OrdinalIgnoreCase) && !IsVowel(className[className.Length - 2]))
        {
            return className.Substring(0, className.Length - 1) + "ies";
        }
        else if (className.EndsWith("s", StringComparison.OrdinalIgnoreCase) ||
                 className.EndsWith("sh", StringComparison.OrdinalIgnoreCase) ||
                 className.EndsWith("ch", StringComparison.OrdinalIgnoreCase) ||
                 className.EndsWith("x", StringComparison.OrdinalIgnoreCase) ||
                 className.EndsWith("z", StringComparison.OrdinalIgnoreCase) ||
                 className.EndsWith("o", StringComparison.OrdinalIgnoreCase))
        {
            return className + "es";
        }
        else
        {
            return className + "s";
        }
    }

    private static bool IsVowel(char c)
    {
        return "aeiou".IndexOf(char.ToLower(c)) >= 0;
    }
}
