using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ListMakerTwo
{
    public class CsvFileSourceReader : ISourceReader
    {
        private IList<LugBulkBuyer> _buyers;
        private IList<LugBulkElement> _elements;
        private IList<LugBulkReservation> _orders;

        public CsvFileSourceReader(List<string> orders, List<string> elements)
        {
            ParseElements(elements);
            ParseBuyers(orders);
            ParseOrders(orders);
        }

        private void ParseBuyers(List<string> orders)
        {
            var result = new List<string>();

            foreach (var line in orders)
            {
                var parts = line.Split(',');

                if (parts[0] == "Nickname")
                    continue;

                if(!result.Any(x => x == parts[0]))
                {
                    result.Add(parts[0]);
                }
            }

            _buyers = new List<LugBulkBuyer>();

            var count = 100;
            foreach(var item in result.OrderBy(x => x))
            {
                _buyers.Add(new LugBulkBuyer() { Id = count, Name = item });
                count++;
            }
        }

        private void ParseOrders(List<string> orders)
        {
            _orders = new List<LugBulkReservation>();

            foreach (var line in orders)
            {
                var parts = line.Split(',');

                if (parts[0] == "Nickname")
                    continue;

                if (parts.Length != 4)
                    throw new Exception("Line doesn't have 4 parts: " + line);

                var amount = 0; 
                
                if(!int.TryParse(parts[3], out amount))
                {
                    throw new Exception("Line missing valid amount: " + line);
                }

                var buyer = _buyers.FirstOrDefault(x => x.Name == parts[0]);

                if(buyer == null)
                {
                    throw new Exception("Line missing valid buyer: " + line);
                }

                var element = _elements.FirstOrDefault(x => x.ElementID == parts[2]);

                if (element == null)
                {
                    throw new Exception("Line missing valid element: " + line);
                }

                var new_order = new LugBulkReservation()
                {
                    Element = element,
                    Buyer = buyer,
                    Amount = amount,
                };

                _orders.Add(new_order);
            }
        }

        private void ParseElements(List<string> elements)
        {
            _elements = new List<LugBulkElement>();

            foreach(var line in elements)
            {
                var parts = line.Split(',');

                if(parts.Length != 5)
                    throw new Exception("Line doesn't have 5 parts: " + line);

                var new_element = new LugBulkElement()
                {
                    BricklinkColor = parts[2],
                    BricklinkDescription = parts[1],
                    BricklinkId = parts[4],
                    ElementID = parts[0],
                    MaterialColor = parts[3],
                };

                _elements.Add(new_element);
            }
        }

        public IList<LugBulkBuyer> GetBuyers()
        {
            return _buyers;
        }

        public IList<LugBulkElement> GetElements()
        {
            return _elements;
        }

        public IList<LugBulkReservation> GetReservations()
        {
            return _orders;
        }
    }
}
