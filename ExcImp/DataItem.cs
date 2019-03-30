using System.Collections.Generic;
using CSharpFunctionalExtensions;

namespace ExcImp
{
    public class DataItem: ValueObject
    {
        public string Code { get; set; }
        public string Number { get; set; }
        public string Vin { get; set; }

        #region Overrides of ValueObject

        protected override IEnumerable<object> GetEqualityComponents()
        {
            yield return Code;
            yield return Number;
            yield return Vin;
        }

        #endregion
    }
}
