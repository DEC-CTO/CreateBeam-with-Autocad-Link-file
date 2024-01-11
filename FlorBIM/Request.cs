using System;
using System.Threading;

namespace FlorBIM
{
    public enum RequestId : int
    {
        None = 0,
        count = 1,
        CreateBeam =2,
        CreateDeck = 3,
        changename = 4,
        gang = 5

    }

    public class Request
    {
        private int m_request = (int)RequestId.None;

        public RequestId Take()
        {
            return (RequestId)Interlocked.Exchange(ref m_request, (int)RequestId.None);
        }

        public void Make(RequestId request)
        {
            Interlocked.Exchange(ref m_request, (int)request);
        }
    }
}
