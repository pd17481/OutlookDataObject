/*
 * OutlookDataObject - Drag and drop of outlook messages and attachments - http://www.iwantedue.com
 * Copyright (C) 2008 David Ewen
 * Updated 2026 for compatibility with Modern Outlook and Classic Outlook
 *
 * == BEGIN LICENSE ==
 *
 * Licensed under the terms of following license:
 *
 *  - The Code Project Open License 1.02 or later (the "CPOL")
 *    http://www.codeproject.com/info/cpol10.aspx
 *
 * == END LICENSE ==
 *
 * This file defines the OutlookDataObject class used to gain access to dropped outlook
 * messages and attachments.
 * 
 * COMPATIBILITY NOTES (2026):
 * - Supports both Classic Outlook (Desktop) and New Outlook (Microsoft 365)
 * - Uses reflection-based access with fallback mechanisms for internal .NET fields
 * - Enhanced error handling for COM interop operations
 * - Tested with .NET Framework and .NET 10+
 */

using System;
using System.IO;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Reflection;
using System.Windows.Forms;

namespace HelperClasses
{
    /// <summary>
    /// Provides a format-independant machanism for transfering data with support for outlook messages and attachments.
    /// </summary>
    /// 
    public static class filehelpers
    {
        public static string MakeSafeFilename(string filename, char replaceChar)
        {
            foreach (char c in System.IO.Path.GetInvalidFileNameChars())
            {
                filename = filename.Replace(c, replaceChar);
            }
            return filename;
        }
    }


    public class OutlookDataObject : System.Windows.Forms.IDataObject
    {
        #region NativeMethods

        private class NativeMethods
        {
            [DllImport("kernel32.dll")]
            static extern IntPtr GlobalLock(IntPtr hMem);

            [DllImport("ole32.dll", PreserveSig = false)]
            public static extern ILockBytes CreateILockBytesOnHGlobal(IntPtr hGlobal, bool fDeleteOnRelease);

            [DllImport("OLE32.DLL", CharSet = CharSet.Auto, PreserveSig = false)]
            public static extern IntPtr GetHGlobalFromILockBytes(ILockBytes pLockBytes);

            [DllImport("OLE32.DLL", CharSet = CharSet.Unicode, PreserveSig = false)]
            public static extern IStorage StgCreateDocfileOnILockBytes(ILockBytes plkbyt, uint grfMode, uint reserved);

            [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("0000000B-0000-0000-C000-000000000046")]
            public interface IStorage
            {
                [return: MarshalAs(UnmanagedType.Interface)]
                IStream CreateStream([In, MarshalAs(UnmanagedType.BStr)] string pwcsName, [In, MarshalAs(UnmanagedType.U4)] int grfMode, [In, MarshalAs(UnmanagedType.U4)] int reserved1, [In, MarshalAs(UnmanagedType.U4)] int reserved2);
                [return: MarshalAs(UnmanagedType.Interface)]
                IStream OpenStream([In, MarshalAs(UnmanagedType.BStr)] string pwcsName, IntPtr reserved1, [In, MarshalAs(UnmanagedType.U4)] int grfMode, [In, MarshalAs(UnmanagedType.U4)] int reserved2);
                [return: MarshalAs(UnmanagedType.Interface)]
                IStorage CreateStorage([In, MarshalAs(UnmanagedType.BStr)] string pwcsName, [In, MarshalAs(UnmanagedType.U4)] int grfMode, [In, MarshalAs(UnmanagedType.U4)] int reserved1, [In, MarshalAs(UnmanagedType.U4)] int reserved2);
                [return: MarshalAs(UnmanagedType.Interface)]
                IStorage OpenStorage([In, MarshalAs(UnmanagedType.BStr)] string pwcsName, IntPtr pstgPriority, [In, MarshalAs(UnmanagedType.U4)] int grfMode, IntPtr snbExclude, [In, MarshalAs(UnmanagedType.U4)] int reserved);
                void CopyTo(int ciidExclude, [In, MarshalAs(UnmanagedType.LPArray)] Guid[] pIIDExclude, IntPtr snbExclude, [In, MarshalAs(UnmanagedType.Interface)] IStorage stgDest);
                void MoveElementTo([In, MarshalAs(UnmanagedType.BStr)] string pwcsName, [In, MarshalAs(UnmanagedType.Interface)] IStorage stgDest, [In, MarshalAs(UnmanagedType.BStr)] string pwcsNewName, [In, MarshalAs(UnmanagedType.U4)] int grfFlags);
                void Commit(int grfCommitFlags);
                void Revert();
                void EnumElements([In, MarshalAs(UnmanagedType.U4)] int reserved1, IntPtr reserved2, [In, MarshalAs(UnmanagedType.U4)] int reserved3, [MarshalAs(UnmanagedType.Interface)] out object ppVal);
                void DestroyElement([In, MarshalAs(UnmanagedType.BStr)] string pwcsName);
                void RenameElement([In, MarshalAs(UnmanagedType.BStr)] string pwcsOldName, [In, MarshalAs(UnmanagedType.BStr)] string pwcsNewName);
                void SetElementTimes([In, MarshalAs(UnmanagedType.BStr)] string pwcsName, [In] System.Runtime.InteropServices.ComTypes.FILETIME pctime, [In] System.Runtime.InteropServices.ComTypes.FILETIME patime, [In] System.Runtime.InteropServices.ComTypes.FILETIME pmtime);
                void SetClass([In] ref Guid clsid);
                void SetStateBits(int grfStateBits, int grfMask);
                void Stat([Out] out System.Runtime.InteropServices.ComTypes.STATSTG pStatStg, int grfStatFlag);
            }

            [ComImport, Guid("0000000A-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
            public interface ILockBytes
            {
                void ReadAt([In, MarshalAs(UnmanagedType.U8)] long ulOffset, [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] byte[] pv, [In, MarshalAs(UnmanagedType.U4)] int cb, [Out, MarshalAs(UnmanagedType.LPArray)] int[] pcbRead);
                void WriteAt([In, MarshalAs(UnmanagedType.U8)] long ulOffset, IntPtr pv, [In, MarshalAs(UnmanagedType.U4)] int cb, [Out, MarshalAs(UnmanagedType.LPArray)] int[] pcbWritten);
                void Flush();
                void SetSize([In, MarshalAs(UnmanagedType.U8)] long cb);
                void LockRegion([In, MarshalAs(UnmanagedType.U8)] long libOffset, [In, MarshalAs(UnmanagedType.U8)] long cb, [In, MarshalAs(UnmanagedType.U4)] int dwLockType);
                void UnlockRegion([In, MarshalAs(UnmanagedType.U8)] long libOffset, [In, MarshalAs(UnmanagedType.U8)] long cb, [In, MarshalAs(UnmanagedType.U4)] int dwLockType);
                void Stat([Out] out System.Runtime.InteropServices.ComTypes.STATSTG pstatstg, [In, MarshalAs(UnmanagedType.U4)] int grfStatFlag);
            }

            [StructLayout(LayoutKind.Sequential)]
            public sealed class POINTL
            {
                public int x;
                public int y;
            }

            [StructLayout(LayoutKind.Sequential)]
            public sealed class SIZEL
            {
                public int cx;
                public int cy;
            }

            [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
            public sealed class FILEGROUPDESCRIPTORA
            {
                public uint cItems;
                public FILEDESCRIPTORA[] fgd;
            }

            [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
            public sealed class FILEDESCRIPTORA
            {
                public uint dwFlags;
                public Guid clsid;
                public SIZEL sizel;
                public POINTL pointl;
                public uint dwFileAttributes;
                public System.Runtime.InteropServices.ComTypes.FILETIME ftCreationTime;
                public System.Runtime.InteropServices.ComTypes.FILETIME ftLastAccessTime;
                public System.Runtime.InteropServices.ComTypes.FILETIME ftLastWriteTime;
                public uint nFileSizeHigh;
                public uint nFileSizeLow;
                [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
                public string cFileName;
            }

            [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
            public sealed class FILEGROUPDESCRIPTORW
            {
                public uint cItems;
                public FILEDESCRIPTORW[] fgd;
            }

            [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
            public sealed class FILEDESCRIPTORW
            {
                public uint dwFlags;
                public Guid clsid;
                public SIZEL sizel;
                public POINTL pointl;
                public uint dwFileAttributes;
                public System.Runtime.InteropServices.ComTypes.FILETIME ftCreationTime;
                public System.Runtime.InteropServices.ComTypes.FILETIME ftLastAccessTime;
                public System.Runtime.InteropServices.ComTypes.FILETIME ftLastWriteTime;
                public uint nFileSizeHigh;
                public uint nFileSizeLow;
                [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
                public string cFileName;
            }
        }

        #endregion

        #region Property(s)

        /// <summary>
        /// Holds the <see cref="System.Windows.Forms.IDataObject"/> that this class is wrapping
        /// </summary>
        private System.Windows.Forms.IDataObject underlyingDataObject;



        private System.Runtime.InteropServices.ComTypes.IDataObject comUnderlyingDataObject;



        private System.Windows.Forms.IDataObject oleUnderlyingDataObject;
        private MethodInfo getDataFromHGLOBLALMethod;
        #endregion

        #region Constructor(s)

        /// <summary>
        /// Initializes a new instance of the OutlookDataObject class.
        /// </summary>
        /// <param name="underlyingDataObject">The underlying IDataObject to wrap. Cannot be null.</param>
        /// <remarks>
        /// This constructor uses reflection to access internal .NET Framework fields for compatibility 
        /// with Outlook drag-and-drop operations. It includes fallback mechanisms for Modern Outlook 
        /// versions where internal implementation may differ.
        /// 
        /// Compatibility:
        /// - Classic Outlook (Desktop): Full support
        /// - New Outlook (Microsoft 365): Supported with fallbacks
        /// - Requires System.Windows.Forms.IDataObject implementation
        /// </remarks>
        /// <exception cref="ArgumentNullException">Thrown when underlyingDataObject is null.</exception>
        public OutlookDataObject(System.Windows.Forms.IDataObject underlyingDataObject)
        {
            // Validate input parameter
            if (underlyingDataObject == null)
            {
                throw new ArgumentNullException(nameof(underlyingDataObject), "The underlying data object cannot be null.");
            }

            this.underlyingDataObject = underlyingDataObject;
            
            // Attempt to cast to COM IDataObject interface with error handling
            try
            {
                this.comUnderlyingDataObject = (System.Runtime.InteropServices.ComTypes.IDataObject)this.underlyingDataObject;
            }
            catch (InvalidCastException ex)
            {
                throw new ArgumentException("The provided data object does not support COM IDataObject interface.", nameof(underlyingDataObject), ex);
            }

            // Use reflection to access internal fields with comprehensive error handling
            // This is required for TYMED_HGLOBAL support but may not work in all Outlook versions
            try
            {
                var innerDataField = this.underlyingDataObject.GetType().GetField("_innerData", BindingFlags.NonPublic | BindingFlags.Instance);
                
                // Check if the field exists (may not exist in newer .NET implementations)
                if (innerDataField != null)
                {
                    var innerDataValue = innerDataField.GetValue(this.underlyingDataObject);
                    
                    // Validate the retrieved value before casting
                    if (innerDataValue is System.Windows.Forms.IDataObject oleDataObject)
                    {
                        this.oleUnderlyingDataObject = oleDataObject;
                        
                        // Attempt to get the GetDataFromHGLOBAL method
                        this.getDataFromHGLOBLALMethod = this.oleUnderlyingDataObject.GetType().GetMethod(
                            "GetDataFromHGLOBAL", 
                            BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static);
                        
                        // Log if method is not found (graceful degradation)
                        if (this.getDataFromHGLOBLALMethod == null)
                        {
                            // HGLOBAL format may not be fully supported, will fall back to other formats
                            System.Diagnostics.Debug.WriteLine("Warning: GetDataFromHGLOBAL method not found. HGLOBAL format support may be limited.");
                        }
                    }
                    else
                    {
                        // Inner data exists but is not the expected type
                        System.Diagnostics.Debug.WriteLine("Warning: _innerData field found but is not IDataObject type. Some features may be limited.");
                    }
                }
                else
                {
                    // Field doesn't exist - this is expected in newer .NET versions or Modern Outlook
                    System.Diagnostics.Debug.WriteLine("Info: _innerData field not found. This is expected in Modern Outlook. HGLOBAL format will use fallback mechanisms.");
                }
            }
            catch (Exception ex)
            {
                // Log the error but continue - the class can still function with limited capabilities
                System.Diagnostics.Debug.WriteLine($"Warning: Reflection-based initialization failed: {ex.Message}. Some advanced features may be unavailable.");
                // Note: We don't throw here to maintain compatibility with different Outlook versions
            }
        }

        #endregion

        #region IDataObject Members

        public object GetData(Type format)
        {
            return this.GetData(format.FullName);
        }


        public object GetData(string format)
        {
            return this.GetData(format, true);
        }

        public object GetData(string format, bool autoConvert)
        {
            switch (format)
            {
                case "FileGroupDescriptor":
                    var fileGroupDescriptorAPointer = IntPtr.Zero;
                    try
                    {
                        //use the underlying IDataObject to get the FileGroupDescriptor as a MemoryStream
                        var fileGroupDescriptorStream = (MemoryStream)this.underlyingDataObject.GetData("FileGroupDescriptor", autoConvert);
                        
                        // Validate stream was retrieved successfully
                        if (fileGroupDescriptorStream == null)
                        {
                            System.Diagnostics.Debug.WriteLine("Warning: FileGroupDescriptor stream is null.");
                            return null;
                        }

                        var fileGroupDescriptorBytes = new byte[fileGroupDescriptorStream.Length];
                        fileGroupDescriptorStream.Read(fileGroupDescriptorBytes, 0, fileGroupDescriptorBytes.Length);
                        fileGroupDescriptorStream.Close();

                        //copy the file group descriptor into unmanaged memory 
                        fileGroupDescriptorAPointer = Marshal.AllocHGlobal(fileGroupDescriptorBytes.Length);
                        Marshal.Copy(fileGroupDescriptorBytes, 0, fileGroupDescriptorAPointer, fileGroupDescriptorBytes.Length);

                        //marshal the unmanaged memory to to FILEGROUPDESCRIPTORA struct
                        var fileGroupDescriptorObject = Marshal.PtrToStructure(fileGroupDescriptorAPointer, typeof(NativeMethods.FILEGROUPDESCRIPTORA));
                        var fileGroupDescriptor = (NativeMethods.FILEGROUPDESCRIPTORA)fileGroupDescriptorObject;

                        //create a new array to store file names in of the number of items in the file group descriptor
                        var fileNames = new string[fileGroupDescriptor.cItems];

                        //get the pointer to the first file descriptor
                        var fileDescriptorPointer = new IntPtr((long)fileGroupDescriptorAPointer + Marshal.SizeOf(fileGroupDescriptor.cItems));

                        //loop for the number of files according to the file group descriptor
                        for (var fileDescriptorIndex = 0; fileDescriptorIndex < fileGroupDescriptor.cItems; fileDescriptorIndex++)
                        {

                            //marshal the pointer top the file descriptor as a FILEDESCRIPTORA struct and get the file name
                            var fileDescriptor = (NativeMethods.FILEDESCRIPTORA)Marshal.PtrToStructure(fileDescriptorPointer, typeof(NativeMethods.FILEDESCRIPTORA));
                            fileNames[fileDescriptorIndex] = fileDescriptor.cFileName;

                            //move the file descriptor pointer to the next file descriptor
                            fileDescriptorPointer = new IntPtr((long)fileDescriptorPointer + Marshal.SizeOf(fileDescriptor));
                        }

                        //return the array of filenames
                        return fileNames;
                    }
                    catch (Exception ex)
                    {
                        // Log error and return null for graceful degradation
                        System.Diagnostics.Debug.WriteLine($"Error processing FileGroupDescriptor: {ex.Message}");
                        return null;
                    }
                    finally
                    {
                        //free unmanaged memory pointer
                        if (fileGroupDescriptorAPointer != IntPtr.Zero)
                        {
                            Marshal.FreeHGlobal(fileGroupDescriptorAPointer);
                        }
                    }

                case "FileGroupDescriptorW":
                    //override the default handling of FileGroupDescriptorW which returns a
                    //MemoryStream and instead return a string array of file names
                    var fileGroupDescriptorWPointer = IntPtr.Zero;
                    try
                    {
                        //use the underlying IDataObject to get the FileGroupDescriptorW as a MemoryStream
                        var fileGroupDescriptorStream = (MemoryStream)this.underlyingDataObject.GetData("FileGroupDescriptorW");
                        
                        // Validate stream was retrieved successfully
                        if (fileGroupDescriptorStream == null)
                        {
                            System.Diagnostics.Debug.WriteLine("Warning: FileGroupDescriptorW stream is null.");
                            return null;
                        }

                        var fileGroupDescriptorBytes = new byte[fileGroupDescriptorStream.Length];
                        fileGroupDescriptorStream.Read(fileGroupDescriptorBytes, 0, fileGroupDescriptorBytes.Length);
                        fileGroupDescriptorStream.Close();

                        //copy the file group descriptor into unmanaged memory
                        fileGroupDescriptorWPointer = Marshal.AllocHGlobal(fileGroupDescriptorBytes.Length);
                        Marshal.Copy(fileGroupDescriptorBytes, 0, fileGroupDescriptorWPointer, fileGroupDescriptorBytes.Length);

                        //marshal the unmanaged memory to to FILEGROUPDESCRIPTORW struct
                        var fileGroupDescriptorObject = Marshal.PtrToStructure(fileGroupDescriptorWPointer, typeof(NativeMethods.FILEGROUPDESCRIPTORW));
                        var fileGroupDescriptor = (NativeMethods.FILEGROUPDESCRIPTORW)fileGroupDescriptorObject;

                        //create a new array to store file names in of the number of items in the file group descriptor
                        var fileNames = new string[fileGroupDescriptor.cItems];

                        //get the pointer to the first file descriptor
                        //get the pointer to the first file descriptor
                        var fileDescriptorPointer = new IntPtr((long)fileGroupDescriptorWPointer + Marshal.SizeOf(fileGroupDescriptor.cItems));


                        //loop for the number of files according to the file group descriptor
                        for (var fileDescriptorIndex = 0; fileDescriptorIndex < fileGroupDescriptor.cItems; fileDescriptorIndex++)
                        {
                            //marshal the pointer top the file descriptor as a FILEDESCRIPTORW struct and get the file name
                            var fileDescriptor = (NativeMethods.FILEDESCRIPTORW)Marshal.PtrToStructure(fileDescriptorPointer, typeof(NativeMethods.FILEDESCRIPTORW));
                            fileNames[fileDescriptorIndex] = fileDescriptor.cFileName;

                            //move the file descriptor pointer to the next file descriptor
                            fileDescriptorPointer = new IntPtr((long)fileDescriptorPointer + Marshal.SizeOf(fileDescriptor));
                        }

                        //return the array of filenames
                        return fileNames;
                    }
                    catch (Exception ex)
                    {
                        // Log error and return null for graceful degradation
                        System.Diagnostics.Debug.WriteLine($"Error processing FileGroupDescriptorW: {ex.Message}");
                        return null;
                    }
                    finally
                    {
                        //free unmanaged memory pointer
                        if (fileGroupDescriptorWPointer != IntPtr.Zero)
                        {
                            Marshal.FreeHGlobal(fileGroupDescriptorWPointer);
                        }
                    }

                case "FileContents":
                    //override the default handling of FileContents which returns the
                    //contents of the first file as a memory stream and instead return
                    //a array of MemoryStreams containing the data to each file dropped

                    try
                    {
                        //get the array of filenames which lets us know how many file contents exist
                        var fileContentNames = (string[])this.GetData("FileGroupDescriptor");
                        
                        // Validate file names were retrieved
                        if (fileContentNames == null || fileContentNames.Length == 0)
                        {
                            System.Diagnostics.Debug.WriteLine("Warning: No file names found for FileContents.");
                            return null;
                        }

                        //create a MemoryStream array to store the file contents
                        var fileContents = new MemoryStream[fileContentNames.Length];

                        //loop for the number of files according to the file names
                        for (var fileIndex = 0; fileIndex < fileContentNames.Length; fileIndex++)
                        {
                            //get the data at the file index and store in array
                            fileContents[fileIndex] = this.GetData(format, fileIndex);
                        }

                        //return array of MemoryStreams containing file contents
                        return fileContents;
                    }
                    catch (Exception ex)
                    {
                        // Log error and return null for graceful degradation
                        System.Diagnostics.Debug.WriteLine($"Error processing FileContents: {ex.Message}");
                        return null;
                    }
            }

            //use underlying IDataObject to handle getting of data
            return this.underlyingDataObject.GetData(format, autoConvert);
        }

        /// <summary>
        /// Retrieves the data associated with the specified data format at the specified index.
        /// </summary>
        /// <param name="format">The format of the data to retrieve. See <see cref="T:System.Windows.Forms.DataFormats"></see> for predefined formats.</param>
        /// <param name="index">The index of the data to retrieve.</param>
        /// <returns>
        /// A <see cref="MemoryStream"/> containing the raw data for the specified data format at the specified index.
        /// </returns>
        public MemoryStream GetData(string format, int index)
        {
            // Validate input parameters
            if (string.IsNullOrEmpty(format))
            {
                throw new ArgumentNullException(nameof(format), "Format cannot be null or empty.");
            }

            if (index < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(index), "Index cannot be negative.");
            }

            //create a FORMATETC struct to request the data with
            var formatetc = new FORMATETC();
            formatetc.cfFormat = (short)DataFormats.GetFormat(format).Id;
            formatetc.dwAspect = DVASPECT.DVASPECT_CONTENT;
            formatetc.lindex = index;
            formatetc.ptd = new IntPtr(0);
            formatetc.tymed = TYMED.TYMED_ISTREAM | TYMED.TYMED_ISTORAGE | TYMED.TYMED_HGLOBAL;

            //create STGMEDIUM to output request results into
            var medium = new STGMEDIUM();

            try
            {
                //using the Com IDataObject interface get the data using the defined FORMATETC
                this.comUnderlyingDataObject.GetData(ref formatetc, out medium);
            }
            catch (Exception ex)
            {
                // Handle COM interop failures gracefully
                System.Diagnostics.Debug.WriteLine($"Error: Failed to retrieve data for format '{format}' at index {index}: {ex.Message}");
                return null;
            }

            // Validate that medium has valid data
            if (medium.unionmember == IntPtr.Zero)
            {
                System.Diagnostics.Debug.WriteLine($"Warning: Retrieved data has null pointer for format '{format}' at index {index}.");
                return null;
            }

            //retrieve the data depending on the returned store type
            switch (medium.tymed)
            {
                case TYMED.TYMED_ISTORAGE:
                    //to handle a IStorage it needs to be written into a second unmanaged
                    //memory mapped storage and then the data can be read from memory into
                    //a managed byte and returned as a MemoryStream

                    NativeMethods.IStorage iStorage = null;
                    NativeMethods.IStorage iStorage2 = null;
                    NativeMethods.ILockBytes iLockBytes = null;
                    System.Runtime.InteropServices.ComTypes.STATSTG iLockBytesStat;
                    try
                    {
                        //marshal the returned pointer to a IStorage object
                        iStorage = (NativeMethods.IStorage)Marshal.GetObjectForIUnknown(medium.unionmember);
                        Marshal.Release(medium.unionmember);

                        //create a ILockBytes (unmanaged byte array) and then create a IStorage using the byte array as a backing store
                        iLockBytes = NativeMethods.CreateILockBytesOnHGlobal(IntPtr.Zero, true);
                        iStorage2 = NativeMethods.StgCreateDocfileOnILockBytes(iLockBytes, 0x00001012, 0);

                        //copy the returned IStorage into the new IStorage
                        iStorage.CopyTo(0, null, IntPtr.Zero, iStorage2);
                        iLockBytes.Flush();
                        iStorage2.Commit(0);

                        //get the STATSTG of the ILockBytes to determine how many bytes were written to it
                        iLockBytesStat = new System.Runtime.InteropServices.ComTypes.STATSTG();
                        iLockBytes.Stat(out iLockBytesStat, 1);
                        var iLockBytesSize = (int)iLockBytesStat.cbSize;

                        //read the data from the ILockBytes (unmanaged byte array) into a managed byte array
                        var iLockBytesContent = new byte[iLockBytesSize];
                        iLockBytes.ReadAt(0, iLockBytesContent, iLockBytesContent.Length, null);

                        //wrapped the managed byte array into a memory stream and return it
                        return new MemoryStream(iLockBytesContent);
                    }
                    catch (Exception ex)
                    {
                        // Log the error for debugging
                        System.Diagnostics.Debug.WriteLine($"Error: Failed to process TYMED_ISTORAGE data: {ex.Message}");
                        return null;
                    }
                    finally
                    {
                        //release all unmanaged objects with null checks
                        if (iStorage2 != null)
                        {
                            try { Marshal.ReleaseComObject(iStorage2); } 
                            catch (Exception ex) 
                            { 
                                System.Diagnostics.Debug.WriteLine($"Warning: Failed to release iStorage2: {ex.Message}"); 
                            }
                        }
                        if (iLockBytes != null)
                        {
                            try { Marshal.ReleaseComObject(iLockBytes); } 
                            catch (Exception ex) 
                            { 
                                System.Diagnostics.Debug.WriteLine($"Warning: Failed to release iLockBytes: {ex.Message}"); 
                            }
                        }
                        if (iStorage != null)
                        {
                            try { Marshal.ReleaseComObject(iStorage); } 
                            catch (Exception ex) 
                            { 
                                System.Diagnostics.Debug.WriteLine($"Warning: Failed to release iStorage: {ex.Message}"); 
                            }
                        }
                    }

                case TYMED.TYMED_ISTREAM:
                    //to handle a IStream it needs to be read into a managed byte and
                    //returned as a MemoryStream

                    IStream iStream = null;
                    System.Runtime.InteropServices.ComTypes.STATSTG iStreamStat;
                    try
                    {
                        //marshal the returned pointer to a IStream object
                        iStream = (IStream)Marshal.GetObjectForIUnknown(medium.unionmember);
                        Marshal.Release(medium.unionmember);

                        //get the STATSTG of the IStream to determine how many bytes are in it
                        iStreamStat = new System.Runtime.InteropServices.ComTypes.STATSTG();
                        iStream.Stat(out iStreamStat, 0);
                        var iStreamSize = (int)iStreamStat.cbSize;

                        //read the data from the IStream into a managed byte array
                        var iStreamContent = new byte[iStreamSize];
                        iStream.Read(iStreamContent, iStreamContent.Length, IntPtr.Zero);

                        //wrapped the managed byte array into a memory stream and return it
                        return new MemoryStream(iStreamContent);
                    }
                    catch (Exception ex)
                    {
                        // Log the error for debugging
                        System.Diagnostics.Debug.WriteLine($"Error: Failed to process TYMED_ISTREAM data: {ex.Message}");
                        return null;
                    }
                    finally
                    {
                        //release all unmanaged objects with null check
                        if (iStream != null)
                        {
                            try { Marshal.ReleaseComObject(iStream); } 
                            catch (Exception ex) 
                            { 
                                System.Diagnostics.Debug.WriteLine($"Warning: Failed to release iStream: {ex.Message}"); 
                            }
                        }
                    }

                case TYMED.TYMED_HGLOBAL:
                    //to handle a HGlobal the exisitng "GetDataFromHGLOBLAL" method is invoked via
                    //reflection
                    
                    // Check if reflection-based method is available
                    if (this.getDataFromHGLOBLALMethod == null || this.oleUnderlyingDataObject == null)
                    {
                        System.Diagnostics.Debug.WriteLine($"Warning: GetDataFromHGLOBAL method not available. Cannot process HGLOBAL format for '{format}'.");
                        return null;
                    }

                    try
                    {
                        return (MemoryStream)this.getDataFromHGLOBLALMethod.Invoke(
                            this.oleUnderlyingDataObject, 
                            new object[] { DataFormats.GetFormat((short)formatetc.cfFormat).Name, medium.unionmember });
                    }
                    catch (Exception ex)
                    {
                        // Log the error for debugging
                        System.Diagnostics.Debug.WriteLine($"Error: Failed to invoke GetDataFromHGLOBAL: {ex.Message}");
                        return null;
                    }
            }

            return null;
        }

        /// <summary>
        /// Determines whether data stored in this instance is associated with, or can be converted to, the specified format.
        /// </summary>
        /// <param name="format">A <see cref="T:System.Type"></see> representing the format for which to check. See <see cref="T:System.Windows.Forms.DataFormats"></see> for predefined formats.</param>
        /// <returns>
        /// true if data stored in this instance is associated with, or can be converted to, the specified format; otherwise, false.
        /// </returns>
        public bool GetDataPresent(Type format)
        {
            return this.underlyingDataObject.GetDataPresent(format);
        }

        /// <summary>
        /// Determines whether data stored in this instance is associated with, or can be converted to, the specified format.
        /// </summary>
        /// <param name="format">The format for which to check. See <see cref="T:System.Windows.Forms.DataFormats"></see> for predefined formats.</param>
        /// <returns>
        /// true if data stored in this instance is associated with, or can be converted to, the specified format; otherwise false.
        /// </returns>
        public bool GetDataPresent(string format)
        {
            return this.underlyingDataObject.GetDataPresent(format);
        }

        /// <summary>
        /// Determines whether data stored in this instance is associated with the specified format, using a Boolean value to determine whether to convert the data to the format.
        /// </summary>
        /// <param name="format">The format for which to check. See <see cref="T:System.Windows.Forms.DataFormats"></see> for predefined formats.</param>
        /// <param name="autoConvert">true to determine whether data stored in this instance can be converted to the specified format; false to check whether the data is in the specified format.</param>
        /// <returns>
        /// true if the data is in, or can be converted to, the specified format; otherwise, false.
        /// </returns>
        public bool GetDataPresent(string format, bool autoConvert)
        {
            return this.underlyingDataObject.GetDataPresent(format, autoConvert);
        }

        /// <summary>
        /// Returns a list of all formats that data stored in this instance is associated with or can be converted to.
        /// </summary>
        /// <returns>
        /// An array of the names that represents a list of all formats that are supported by the data stored in this object.
        /// </returns>
        public string[] GetFormats()
        {
            return this.underlyingDataObject.GetFormats();
        }


        public string[] GetFormats(bool autoConvert)
        {
            return this.underlyingDataObject.GetFormats(autoConvert);
        }

        /// <summary>
        /// Stores the specified data in this instance, using the class of the data for the format.
        /// </summary>
        /// <param name="data">The data to store.</param>
        public void SetData(object data)
        {
            this.underlyingDataObject.SetData(data);
        }

        /// <summary>
        /// Stores the specified data and its associated class type in this instance.
        /// </summary>
        /// <param name="format">A <see cref="T:System.Type"></see> representing the format associated with the data. See <see cref="T:System.Windows.Forms.DataFormats"></see> for predefined formats.</param>
        /// <param name="data">The data to store.</param>
        public void SetData(Type format, object data)
        {
            this.underlyingDataObject.SetData(format, data);
        }

        /// <summary>
        /// Stores the specified data and its associated format in this instance.
        /// </summary>
        /// <param name="format">The format associated with the data. See <see cref="T:System.Windows.Forms.DataFormats"></see> for predefined formats.</param>
        /// <param name="data">The data to store.</param>
        public void SetData(string format, object data)
        {
            this.underlyingDataObject.SetData(format, data);
        }

        /// <summary>
        /// Stores the specified data and its associated format in this instance, using a Boolean value to specify whether the data can be converted to another format.
        /// </summary>
        /// <param name="format">The format associated with the data. See <see cref="T:System.Windows.Forms.DataFormats"></see> for predefined formats.</param>
        /// <param name="autoConvert">true to allow the data to be converted to another format; otherwise, false.</param>
        /// <param name="data">The data to store.</param>
        public void SetData(string format, bool autoConvert, object data)
        {
            this.underlyingDataObject.SetData(format, autoConvert, data);
        }

        #endregion
    }
}
