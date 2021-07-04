using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace OnlineCinemaAdminApplication.Models
{
    public class Ticket
    {
        public Guid Id { get; set; }
        public string TicketName { get; set; }
        
        public string TicketImage { get; set; }
       
        public int TicketPrice { get; set; }
      
        public string TicketDescription { get; set; }
       
        public DateTime TicketDate { get; set; }

        public string MovieGenre { get; set; }

        public int MovieRating { get; set; }
    }
}
