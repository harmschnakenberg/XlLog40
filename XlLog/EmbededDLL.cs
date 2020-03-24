using System;
using System.Linq;
using System.Reflection;

namespace Kreutztraeger
{
    // Quelle: https://wojciechkulik.pl/csharp/embedded-class-libraries-dll
    static class EmbededDLL
    {
        private static Assembly ExecutingAssembly = Assembly.GetExecutingAssembly();
        private static string[] EmbeddedLibraries = ExecutingAssembly.GetManifestResourceNames().Where(x => x.EndsWith(".dll")).ToArray();

        /// <summary>
        /// Lädt alle *.dll-Dateien aus Resources (DLL mit Builtvorgang: Eingebettete Resource)   
        /// Damit keine Kopie der DLL erstellt wird: Verweise -> DLL-Name -> Eigenschaften -> Lokale Kopie = false
        /// Nur notwendig, wenn DLLs in diese exe integriert werden sollen.
        /// </summary>
        [STAThread]
        internal static void LoadDlls()
        {
            // Attach custom event handler
            AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;
        }

        private static Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            // Get assembly name
            var assemblyName = new AssemblyName(args.Name).Name + ".dll";

            // Get resource name
            var resourceName = EmbeddedLibraries.FirstOrDefault(x => x.EndsWith(assemblyName));
            if (resourceName == null)
            {
                return null;
            }

            // Load assembly from resource
            using (var stream = ExecutingAssembly.GetManifestResourceStream(resourceName))
            {
                var bytes = new byte[stream.Length];
                stream.Read(bytes, 0, bytes.Length);
                return Assembly.Load(bytes);
            }
        }

    }
}
