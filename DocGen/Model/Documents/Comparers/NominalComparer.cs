using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace DocGen.Model.Documents.Comparers
{
    class NominalComparer : IComparer<Components>
    {
        Dictionary<string, int> ohms = new Dictionary<string, int>()
        {
            { "пОм", 8},
            { "нОм", 7},
            { "мкОм", 6},
            { "мОм", 5},
            { "Ом", 4},
            { "кОм", 3},
            { "МОм", 2},
            { "ГОм", 1}
        };

        Dictionary<string, int> farads = new Dictionary<string, int>()
        {
            { "пФ", 7},
            { "нФ", 6},
            { "мкФ", 5},
            { "мФ", 4},
            { "Ф", 3},
            { "кФ", 2},
            { "МФ", 1},
        };

        private Regex regexFarad = new Regex(@"[а-яА-Я]*Ф", RegexOptions.Compiled);
        private Regex regexOhm = new Regex(@"[а-яА-Я]*Ом", RegexOptions.Compiled);

        // private Regex regexValueFarad = new Regex(@"[0-9.,]+[а-яА-Я ]*Ф", RegexOptions.Compiled);
        // private Regex regexValueOhm = new Regex(@"[0-9.,]+[а-яА-Я ]*Ом", RegexOptions.Compiled);

        private Regex regexValueFarad = new Regex(@"[0-9.,]+?(?=[а-яА-Я ]*Ф)", RegexOptions.Compiled);
        private Regex regexValueOhm = new Regex(@"[0-9.,]+?(?=[а-яА-Я ]*Ом)", RegexOptions.Compiled);

        private Regex regexSeriesFarad =
                new Regex(@"^.*?(?=[0-9.,]+[а-яА-Я ]*Ф)", RegexOptions.Compiled);
        private Regex regexSeriesOhm =
                new Regex(@"^.*?(?=[0-9.,]+[а-яА-Я ]*Ом)", RegexOptions.Compiled);

        private string farad = "Ф";
        private string ohm = "Ом";

        public int Compare(Components c1, Components c2)
        {

            // получить два partnumber
            string pn1 = c1.Part.ManufacturerPartNumber;
            string pn2 = c2.Part.ManufacturerPartNumber;

            string desc1 = c1.Part.Description;
            string desc2 = c2.Part.Description;

            string des1 = c1.GetDesignators();
            string des2 = c2.GetDesignators();

            if (des1.Contains("C") && des2.Contains("C"))
            {
                return DoCompare(c1.Part, c2.Part, farad, regexFarad,
                                regexValueFarad, regexSeriesFarad, farads);
            }

            if (des1.Contains("R") && des2.Contains("R"))
            {
                return DoCompare(c1.Part, c2.Part, ohm, regexOhm,
                                regexValueOhm, regexSeriesOhm, ohms);
            }

            return pn1.CompareTo(pn2);
        }



        private int DoCompare(Part part1, Part part2, string unit,
            Regex regexUnit, Regex regexValueUnit, Regex regexSeriesUnit,
            Dictionary<string, int> units)
        {
            string pn1 = null;
            string pn2 = null;

            // проверить, содержат ли оба номинал
            if (!String.IsNullOrEmpty(part1.ManufacturerPartNumber)
                && !String.IsNullOrEmpty(part2.ManufacturerPartNumber)
                && part1.ManufacturerPartNumber.Contains(unit)
                && part2.ManufacturerPartNumber.Contains(unit))
            {
                pn1 = part1.ManufacturerPartNumber;
                pn2 = part2.ManufacturerPartNumber;
            }
            else if (!String.IsNullOrEmpty(part1.Description)
              && !String.IsNullOrEmpty(part2.Description)
              && part1.Description.Contains(unit)
              && part2.Description.Contains(unit))
            {
                pn1 = part1.Description;
                pn2 = part2.Description;
            }
            else
            {
                return part1.ManufacturerPartNumber.CompareTo(part2.ManufacturerPartNumber);
            }

            // найти части до номинала;
            MatchCollection matchedPrefixPN1 = regexSeriesUnit.Matches(pn1);
            MatchCollection matchedPrefixPN2 = regexSeriesUnit.Matches(pn2);

            string prefixPN1 = matchedPrefixPN1[0].Value;
            string prefixPN2 = matchedPrefixPN2[0].Value;

            // сравнить эти части
            // если части до номинала равны -
            if (prefixPN1.Equals(prefixPN2))
            {
                // найти номиналы
                MatchCollection matchedValuePN1 = regexValueUnit.Matches(pn1);
                MatchCollection matchedValuePN2 = regexValueUnit.Matches(pn2);

                string valuePN1 = matchedValuePN1[0].Value;
                string valuePN2 = matchedValuePN2[0].Value;

                // найти единицы измерения
                MatchCollection matchedUnitPN1 = regexUnit.Matches(pn1);
                MatchCollection matchedUnitPN2 = regexUnit.Matches(pn2);

                string unitPN1 = matchedUnitPN1[0].Value;
                string unitPN2 = matchedUnitPN2[0].Value;

                // сравнить единицы измерения
                // если равны
                if (unitPN1.Equals(unitPN2))
                {
                    // сравнить номиналы


                    decimal value1 = convert(valuePN1);
                    decimal value2 = convert(valuePN2);

                    // вернуть значение
                    return value1.CompareTo(value2);
                }
                // если нет
                else
                {
                    // сравнить единицы измерения                            
                    // вернуть значение
                    return units[unitPN2].CompareTo(units[unitPN1]);
                }

            }
            // если части до номинала отличаются
            else
            {
                // сравнить эти части
                // вернуть значение
                // return prefixPN1.CompareTo(prefixPN2);
                return part1.ManufacturerPartNumber.CompareTo(part2.ManufacturerPartNumber);
            }

        }

        private decimal convert(string stringVal)
        {
            decimal decimalVal = 0;
            try
            {
                decimalVal = System.Convert.ToDecimal(stringVal);
            }
            catch (System.OverflowException)
            {
                decimalVal = 0;
            }
            catch (System.FormatException)
            {
                decimalVal = 0;
            }
            catch (System.ArgumentNullException)
            {
                decimalVal = 0;
            }

            return decimalVal;
        }
    }
}
