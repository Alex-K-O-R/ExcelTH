using System;

namespace TableHandlers
{
    public static class Utilities
    {
        public static T[] add<T>(ref T[] u_where, T u_what)
        {
            if (u_what != null)
            {
                T[] tmp = new T[(u_where == null) ? 1 : (u_where.Length + 1)];
                if ((u_where != null) && (u_where.Length > 0)) u_where.CopyTo(tmp, 0);
                tmp[tmp.Length - 1] = u_what;
                u_where = new T[tmp.Length];
                u_where = tmp;
            }

            return u_where;
        }


    }

    public static class Classes
    {
        public class DTInterval
        {
            public DateTime begin_dt;
            public DateTime end_dt;

            public DTInterval(DateTime? end_dt = null, DateTime ? begin_dt = null)
            {
                this.begin_dt = (DateTime)((begin_dt==null)?DateTime.Now.AddDays(-7):begin_dt);
                this.end_dt = (DateTime)((end_dt == null) ? DateTime.Now : end_dt);
            }

            public void setDepth(int hours)
            {
                this.begin_dt = this.end_dt.AddHours(hours);
            }
        }

    }
}
