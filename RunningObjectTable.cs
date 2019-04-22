using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;

namespace System.Runtime.InteropServices
{
    public static class RunningObjectTable
    {

        [DllImport("ole32.dll")]
        static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);

        [DllImport("ole32.dll")]
        static extern void GetRunningObjectTable(int reserved, out IRunningObjectTable prot);




        [DllImport("ole32.dll", PreserveSig = false)]
        static extern void CLSIDFromProgIDEx([MarshalAs(UnmanagedType.LPWStr)] string progId, out Guid clsid);

        [DllImport("ole32.dll", PreserveSig = false)]
        static extern void CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string progId, out Guid clsid);

        public static T TryGetActiveObject<T>(string ProgId, Func<RunningObjectTableEntry, T, bool> Filter = null, Func<T> Default = null)
        {
            var ret = default(T);

            try
            {
                ret = GetActiveObject<T>(ProgId, Filter, Default);
            }
            catch (Exception ex)
            {
                ex.Ignore();
            }

            return ret;
        }

        public static T GetActiveObject<T>(string ProgId, Func<RunningObjectTableEntry, T, bool> Filter = null, Func<T> Default = null)
        {
            var ret = GetActiveEntries<T>(ProgId, Filter)
                .Select(x => x.Object)
                .OfType<T>()
                .FirstOrDefault()
                ;

            if (object.Equals(ret, default(T)) && Default != null)
            {
                ret = Default();
            }

            return ret;
        }

        public static IEnumerable<Object> GetActiveObjects<T>(string ProgId, Func<RunningObjectTableEntry, T, bool> Filter = null)
        {
            return GetActiveEntries<T>(ProgId, Filter).Select(x => x.Object);
        }

        public static IEnumerable<Object> GetActiveObjects(string ProgId)
        {
            return GetActiveEntries(ProgId).Select(x => x.Object);
        }

        public static IEnumerable<Object> GetActiveObjects()
        {
            return GetActiveEntries().Select(x => x.Object);
        }

        public static RunningObjectTableEntry GetActiveEntry<T>(string ProgId, Func<RunningObjectTableEntry, T, bool> Filter = null)
        {
            return GetActiveEntries<T>(ProgId, Filter).FirstOrDefault();
        }

        public static IEnumerable<RunningObjectTableEntry> GetActiveEntries<T>(string ProgId, Func<RunningObjectTableEntry, T, bool> Filter = null)
        {
            Filter = Filter ?? ((a, b) => true);

            var IE = GetActiveEntries(ProgId);
            foreach (var item in IE)
            {
                var ret = default(RunningObjectTableEntry);
                try
                {
                    if (item.Object is T TItem && Filter(item, TItem))
                    {
                        ret = item;
                    }
                }
                catch (Exception ex)
                {
                    ex.Ignore();
                }

                if (ret != null)
                {
                    yield return item;
                }

            }
        }

        public static IEnumerable<RunningObjectTableEntry> GetActiveEntries(string progID)
        {
            //Guid ClassID = default(Guid);

            var ClassID = Guid.Empty;

            try
            {
                CLSIDFromProgIDEx(progID, out ClassID);
            }
            catch (Exception)
            {
                CLSIDFromProgID(progID, out ClassID);
            }
            var Items = GetActiveEntries().ToList();
            var ret = Items.Where(x =>
                x.ClassID == ClassID ||
                ClassID == null ? false : x.DisplayName.ToLower().Contains(ClassID.ToString().ToLower())
                ).ToList()
                ;
            return ret;

        }

        // Requires Using System.Runtime.InteropServices.ComTypes

        // Get all running instance by querying ROT

        public static IEnumerable<RunningObjectTableEntry> GetActiveEntries()
        {

            // get Running Object Table ...
            GetRunningObjectTable(0, out var Rot);
            if (Rot == null)
            {
                yield break;
            }

            // get enumerator for ROT entries
            Rot.EnumRunning(out var monikerEnumerator);

            if (monikerEnumerator == null)
            {
                yield break;
            }


            monikerEnumerator.Reset();



            var pNumFetched = new IntPtr();

            var monikers = new IMoniker[1];



            // go through all entries and identifies app instances

            while (monikerEnumerator.Next(1, monikers, pNumFetched) == 0)
            {
                var ret = default(RunningObjectTableEntry);

                try
                {

                    CreateBindCtx(0, out var bindCtx);

                    if (bindCtx == null)

                        continue;

                    Rot.GetObject(monikers[0], out var ComObject);
                    monikers[0].GetDisplayName(bindCtx, null, out var displayName);
                    monikers[0].GetClassID(out var ClassId);

                    ret = new RunningObjectTableEntry()
                    {
                        Object = ComObject,
                        DisplayName = displayName,
                        ClassID = ClassId,
                    };
                }
                catch (Exception ex)
                {
                    ex.Ignore();
                }

                if (ret != null)
                {
                    yield return ret;
                }

            }

        }

    }

    [DebuggerDisplay("{ClassID}: {DisplayName}")]
    public class RunningObjectTableEntry
    {
        public string DisplayName { get; set; }
        public Guid ClassID { get; set; }
        public object Object { get; set; }
    }

}
