using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using Task = System.Threading.Tasks.Task;

namespace FixTextDisplay
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [ProvideAutoLoad(UIContextGuids.NoSolution, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideAutoLoad(UIContextGuids.SolutionExists, PackageAutoLoadFlags.BackgroundLoad)]
    [Guid(FixTextDisplayPackage.PackageGuidString)]
    public sealed class FixTextDisplayPackage : AsyncPackage
    {
        /// <summary>
        /// FixTextDisplayPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "0A322AC6-2FFF-4ECF-8083-B509634D4B14";

        const string VS_EDITOR_ASSEMBLY_NAME = "Microsoft.VisualStudio.Platform.VSEditor";
        const string UI_INTERNAL_ASSEMBLY_NAME = "Microsoft.VisualStudio.Shell.UI.Internal";
        const string VS_DEBUG_CORE_UI_ASSEMBLY_NAME = "VSDebugCoreUI";

        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        /// <param name="cancellationToken">A cancellation token to monitor for initialization cancellation, which can occur when VS is shutting down.</param>
        /// <param name="progress">A provider for progress updates.</param>
        /// <returns>A task representing the async work of package initialization, or an already completed task if there is none. Do not return null from this method.</returns>
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            try
            {
                // When initialized asynchronously, the current thread may be a background thread at this point.
                // Do any initialization that requires the UI thread after switching to the UI thread.
                await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
                FixDisplay();
            }
            catch(Exception)
            {
                //TODO make some logging
            }
        }

        #endregion

        private static object CoerceTextFormatting(DependencyObject d, object value)
        {
            return TextFormattingMode.Display;
        }

        static void FixDisplay()
        {
            //All textblocks
            TextOptions.TextFormattingModeProperty.OverrideMetadata(typeof(TextBlock),
                new FrameworkPropertyMetadata(TextFormattingMode.Display, 
                    FrameworkPropertyMetadataOptions.AffectsMeasure | 
                    FrameworkPropertyMetadataOptions.AffectsRender | 
                    FrameworkPropertyMetadataOptions.Inherits, null, CoerceTextFormatting));

            //All textboxes
            TextOptions.TextFormattingModeProperty.OverrideMetadata(typeof(TextBox),
                new FrameworkPropertyMetadata(TextFormattingMode.Display,
                    FrameworkPropertyMetadataOptions.AffectsMeasure |
                    FrameworkPropertyMetadataOptions.AffectsRender |
                    FrameworkPropertyMetadataOptions.Inherits, null, CoerceTextFormatting));

            if(true)
            {
                //Modern exception dialog message
                Style style = new Style(typeof(RichTextBox));
                style.Setters.Add(new Setter(TextOptions.TextFormattingModeProperty, TextFormattingMode.Display));
                Application.Current.Resources.Add(typeof(RichTextBox), style);
            }
            //if (true)
            //{
            //    Style style = new Style(typeof(TextBlock));
            //    style.Setters.Add(new Setter(TextOptions.TextFormattingModeProperty, TextFormattingMode.Display));
            //    Application.Current.Resources.Add(typeof(TextBlock), style);
            //}
            var assemblies = AppDomain.CurrentDomain.GetAssemblies();
            
            if(true)
            {
                var assembly = assemblies.FirstOrDefault(a => a.GetName().Name == VS_EDITOR_ASSEMBLY_NAME);
                if (assembly != null)
                {
                    HandleVsEditorAssembly(assembly);
                }
                else
                {
                    AppDomain.CurrentDomain.AssemblyLoad += CurrentDomain_VSEditorLoad;
                }
            }
            if(true)
            {
                var assembly = assemblies.FirstOrDefault(a => a.GetName().Name == UI_INTERNAL_ASSEMBLY_NAME);
                if (assembly != null)
                {
                    HandleUIInternalAssembly(assembly);
                }
                else
                {
                    AppDomain.CurrentDomain.AssemblyLoad += CurrentDomain_UIInternalLoad;
                }
            }
            if (true)
            {
                var assembly = assemblies.FirstOrDefault(a => a.GetName().Name == VS_DEBUG_CORE_UI_ASSEMBLY_NAME);
                if (assembly != null)
                {
                    HandleVsDebugCoreUIAssembly(assembly);
                }
                else
                {
                    AppDomain.CurrentDomain.AssemblyLoad += CurrentDomain_VsDebugCoreUILoad;
                }
            }
        }

        private static void CurrentDomain_VSEditorLoad(object sender, AssemblyLoadEventArgs args)
        {
            var assembly = args.LoadedAssembly;
            if (assembly.GetName().Name == VS_EDITOR_ASSEMBLY_NAME)
            {
                HandleVsEditorAssembly(assembly);
                AppDomain.CurrentDomain.AssemblyLoad -= CurrentDomain_VSEditorLoad;
            }
        }

        static void HandleVsEditorAssembly(Assembly assembly)
        {
            var types = assembly.GetTypes().ToArray();
            if (true)
            {
                //Test result info panel
                var docType = types.Where(t => t.FullName == "Microsoft.VisualStudio.Text.Editor.Implementation.WpfTextView").FirstOrDefault();
                if (docType != null)
                {
                    TextOptions.TextFormattingModeProperty.OverrideMetadata(docType,
                        new FrameworkPropertyMetadata(TextFormattingMode.Display,
                            FrameworkPropertyMetadataOptions.AffectsMeasure |
                            FrameworkPropertyMetadataOptions.AffectsRender |
                            FrameworkPropertyMetadataOptions.Inherits, null, CoerceTextFormatting));
                }
            }
        }

        private static void CurrentDomain_UIInternalLoad(object sender, AssemblyLoadEventArgs args)
        {
            var assembly = args.LoadedAssembly;
            if (assembly.GetName().Name == UI_INTERNAL_ASSEMBLY_NAME)
            {
                HandleUIInternalAssembly(assembly);
                AppDomain.CurrentDomain.AssemblyLoad -= CurrentDomain_UIInternalLoad;
            }
        }

        static void HandleUIInternalAssembly(Assembly assembly)
        {
            var types = assembly.GetTypes().ToArray();
            if (true)
            {
                //ConfigurationManager
                var docType = types.Where(t => t.FullName == "Microsoft.VisualStudio.PlatformUI.ListViewGrid.ListViewGrid").FirstOrDefault();
                if (docType != null)
                {
                    var style = new Style(docType);
                    style.Setters.Add(new Setter(TextElement.FontSizeProperty, 13.0));
                    Application.Current.Resources.Add(docType, style);
                }
            }
        }

        private static void CurrentDomain_VsDebugCoreUILoad(object sender, AssemblyLoadEventArgs args)
        {
            var assembly = args.LoadedAssembly;
            if (assembly.GetName().Name == VS_DEBUG_CORE_UI_ASSEMBLY_NAME)
            {
                HandleVsDebugCoreUIAssembly(assembly);
                AppDomain.CurrentDomain.AssemblyLoad -= CurrentDomain_VsDebugCoreUILoad;
            }
        }

        static void HandleVsDebugCoreUIAssembly(Assembly assembly)
        {
            var types = assembly.GetTypes().ToArray();
            if (true)
            {
                //ExceptionControl
                var docType = types.Where(t => t.FullName == "VSDebugCoreUI.ExceptionHelper.Controls.ExceptionControl").FirstOrDefault();
                if (docType != null)
                {
                    var style = new Style(docType);
                    style.Setters.Add(new Setter(TextOptions.TextFormattingModeProperty, TextFormattingMode.Display));
                    Application.Current.Resources.Add(docType, style);
                }
            }
        }
    }
}
