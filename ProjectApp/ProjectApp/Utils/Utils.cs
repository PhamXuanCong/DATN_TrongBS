using Autodesk.Revit.DB;
using ProjectApp.ModelFromCad;

namespace ProjectApp.Utils
{
    public static class Utils
    {
        public static XYZ ToXyz(this XyzData data)
        {
            return new XYZ(data.X.MmToFoot(), data.Y.MmToFoot(), data.Z.MmToFoot());
        }

        public static XYZ ToXyzfit(this XYZ data)
        {
            return new XYZ(data.X.MmToFoot(), data.Y.MmToFoot(), data.Z.MmToFoot());
        }

        public static tsource MinBy2<tsource, tkey>(
            this IEnumerable<tsource> source,
            Func<tsource, tkey> selector)
        {
            return source.MinBy2(selector, Comparer<tkey>.Default);
        }

        public static tsource MinBy2<tsource, tkey>(
            this IEnumerable<tsource> source,
            Func<tsource, tkey> selector,
            IComparer<tkey> comparer)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));
            if (selector == null)
                throw new ArgumentNullException(nameof(selector));
            if (comparer == null)
                throw new ArgumentNullException(nameof(comparer));
            using (IEnumerator<tsource> sourceIterator = source.GetEnumerator())
            {
                if (!sourceIterator.MoveNext())
                    throw new InvalidOperationException("Sequence was empty");
                tsource min = sourceIterator.Current;
                tkey minKey = selector(min);
                while (sourceIterator.MoveNext())
                {
                    tsource candidate = sourceIterator.Current;
                    tkey candidateProjected = selector(candidate);
                    if (comparer.Compare(candidateProjected, minKey) < 0)
                    {
                        min = candidate;
                        minKey = candidateProjected;
                    }
                }

                return min;
            }
        }

        public static double MmToFoot(this double mm)
        {
            return mm / 304.79999999999995;
        }

        public static double FootToMm(this double mm)
        {
            return mm * 304.79999999999995;
        }

        public static XYZ EditZ(this XYZ p, double z)
        {
            return new XYZ(p.X, p.Y, z);
        }

        public static Line CreateLineByPointAndDirection(this XYZ p, XYZ direction)
        {
            return Line.CreateBound(p, p.Add(direction));
        }

        public static List<Curve> ToCurves(this CurveArray curveArray)
        {
            var result = new List<Curve>();
            if (curveArray == null) return result;

            foreach (Curve c in curveArray)
            {
                if (c != null) result.Add(c);
            }
            return result;
        }
    }
}
