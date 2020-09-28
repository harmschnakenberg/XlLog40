// This file is used by Code Analysis to maintain SuppressMessage
// attributes that are applied to this project.
// Project-level suppressions either have no target or are given
// a specific target and scoped to a namespace, type, member, etc.

using System.Diagnostics.CodeAnalysis;

[assembly: System.Diagnostics.CodeAnalysis.SuppressMessage("Security", "CA2100:Review SQL queries for security vulnerabilities", Justification = "<Ausstehend>", Scope = "member", Target = "~M:Kreutztraeger.Sql.SqlQyery(System.String,System.String)~System.Data.DataTable")]
[assembly: System.Diagnostics.CodeAnalysis.SuppressMessage("Stil", "IDE1006:Benennungsstile", Justification = "<Ausstehend>", Scope = "member", Target = "~M:Kreutztraeger.NativeMethods.wwHeap_Register(System.Int32,System.Int16@)~System.Boolean")]
[assembly: System.Diagnostics.CodeAnalysis.SuppressMessage("Stil", "IDE1006:Benennungsstile", Justification = "<Ausstehend>", Scope = "member", Target = "~M:Kreutztraeger.NativeMethods.wwHeap_Unregister~System.Boolean")]
[assembly: System.Diagnostics.CodeAnalysis.SuppressMessage("Design", "CA1031:Do not catch general exception types", Justification = "<Ausstehend>", Scope = "member", Target = "~M:Kreutztraeger.Config.CreateConfig(System.String)")]
[assembly: SuppressMessage("Globalization", "CA2101:Specify marshaling for P/Invoke string arguments", Justification = "<Ausstehend>", Scope = "member", Target = "~M:Kreutztraeger.NativeMethods.PtAccHandleCreate(System.Int32,System.String)~System.Int32")]
[assembly: SuppressMessage("Globalization", "CA2101:Specify marshaling for P/Invoke string arguments", Justification = "<Ausstehend>", Scope = "member", Target = "~M:Kreutztraeger.NativeMethods.PtAccActivate(System.Int32,System.String)~System.Int32")]
[assembly: SuppressMessage("Globalization", "CA2101:Specify marshaling for P/Invoke string arguments", Justification = "<Ausstehend>", Scope = "member", Target = "~M:Kreutztraeger.NativeMethods.PtAccReadM(System.Int32,System.Int32,System.Text.StringBuilder,System.Int32)~System.Int32")]
[assembly: SuppressMessage("Globalization", "CA2101:Specify marshaling for P/Invoke string arguments", Justification = "<Ausstehend>", Scope = "member", Target = "~M:Kreutztraeger.NativeMethods.PtAccWriteM(System.Int32,System.Int32,System.String)~System.Int32")]
[assembly: SuppressMessage("Design", "CA1052:Static holder types should be Static or NotInheritable", Justification = "<Ausstehend>", Scope = "type", Target = "~T:Kreutztraeger.SetNewSystemTime")]
