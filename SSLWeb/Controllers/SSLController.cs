  using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

using SSLWeb.Models;

namespace SSLWeb.Controllers
{
    public class SSLController : ApiController
    {
        // GET: api/SSL
        public IEnumerable<SSL> Get()
        {
            return new SSL[] { };
        }

        // GET: api/SSL/5
        public string Get(int id)
        {
            return "Sole Source Letter";
        }

        // POST: api/SSL
        public void Post([FromBody]SSL SoleSourceLetter)
        {
            SoleSourceLetter.CreateLetter();
        }

        // PUT: api/SSL/5
        public void Put(int id, [FromBody]SSL value)
        {
        }

        // DELETE: api/SSL/5
        public void Delete(int id)
        {
        }
    }
}
