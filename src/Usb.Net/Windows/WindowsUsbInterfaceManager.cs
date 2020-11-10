﻿using Device.Net;
using Device.Net.Exceptions;
using Device.Net.Windows;
using Microsoft.Extensions.Logging;
using Microsoft.Win32.SafeHandles;
using System;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace Usb.Net.Windows
{
    public class WindowsUsbInterfaceManager : UsbInterfaceManager, IUsbInterfaceManager
    {
        #region Fields
        private bool disposed;
        private SafeFileHandle _DeviceHandle;
        private SafeFileHandle _InterfaceHandle;
        protected ushort? ReadBufferSizeProtected { get; set; }
        protected ushort? WriteBufferSizeProtected { get; set; }
        #endregion

        #region Public Properties
        public bool IsInitialized => _DeviceHandle != null && !_DeviceHandle.IsInvalid;
        public string DeviceId { get; }

        //TODO: Null checking here. These will error if the device doesn't have a value or it is not initialized
        public ushort WriteBufferSize => WriteBufferSizeProtected ?? WriteUsbInterface.ReadBufferSize;
        public ushort ReadBufferSize => ReadBufferSizeProtected ?? ReadUsbInterface.ReadBufferSize;
        #endregion

        #region Constructor
        public WindowsUsbInterfaceManager(
            string deviceId,
            ILoggerFactory loggerFactory = null,
            ushort? readBufferLength = null,
            ushort? writeBufferLength = null) : base(loggerFactory)
        {
            ReadBufferSizeProtected = readBufferLength;
            WriteBufferSizeProtected = writeBufferLength;
            DeviceId = deviceId;
        }
        #endregion

        #region Private Methods
        private void Initialize()
        {
            using var logScope = Logger.BeginScope("DeviceId: {deviceId} Call: {call}", DeviceId, nameof(Initialize));

            try
            {

                Close();

                int errorCode;

                if (string.IsNullOrEmpty(DeviceId))
                {
                    throw new ValidationException(
                        $"{nameof(ConnectedDeviceDefinition)} must be specified before {nameof(InitializeAsync)} can be called.");
                }

                _DeviceHandle = APICalls.CreateFile(DeviceId,
                    FileAccessRights.GenericWrite | FileAccessRights.GenericRead,
                    APICalls.FileShareRead | APICalls.FileShareWrite, IntPtr.Zero, APICalls.OpenExisting,
                    APICalls.FileAttributeNormal | APICalls.FileFlagOverlapped, IntPtr.Zero);

                if (_DeviceHandle.IsInvalid)
                {
                    //TODO: is error code useful here?
                    errorCode = Marshal.GetLastWin32Error();
                    if (errorCode > 0) throw new ApiException($"Device handle no good. Error code: {errorCode}");
                }

                Logger.LogInformation(Messages.SuccessMessageGotWriteAndReadHandle);

#pragma warning disable CA2000 //We need to hold on to this handle
                var isSuccess = WinUsbApiCalls.WinUsb_Initialize(_DeviceHandle, out _InterfaceHandle);
#pragma warning restore CA2000
                WindowsDeviceBase.HandleError(isSuccess, Messages.ErrorMessageCouldntIntializeDevice);

                var connectedDeviceDefinition = GetDeviceDefinition(_InterfaceHandle, DeviceId, Logger);

                if (!WriteBufferSizeProtected.HasValue)
                {
                    if (!connectedDeviceDefinition.WriteBufferSize.HasValue)
                        throw new ValidationException("Write buffer size not specified");
                    WriteBufferSizeProtected = (ushort)connectedDeviceDefinition.WriteBufferSize.Value;
                }

                if (!ReadBufferSizeProtected.HasValue)
                {
                    if (!connectedDeviceDefinition.ReadBufferSize.HasValue)
                        throw new ValidationException("Read buffer size not specified");
                    ReadBufferSizeProtected = (ushort)connectedDeviceDefinition.ReadBufferSize.Value;
                }

                //Get the first (default) interface
#pragma warning disable CA2000 //Ths should be disposed later
                var defaultInterface = GetInterface(_InterfaceHandle);

                UsbInterfaces.Add(defaultInterface);

                byte i = 0;
                while (true)
                {
                    isSuccess = WinUsbApiCalls.WinUsb_GetAssociatedInterface(_InterfaceHandle, i,
                        out var interfacePointer);
                    if (!isSuccess)
                    {
                        errorCode = Marshal.GetLastWin32Error();
                        if (errorCode == APICalls.ERROR_NO_MORE_ITEMS) break;

                        throw new ApiException(
                            $"Could not enumerate interfaces for device. Error code: {errorCode}");
                    }

                    var associatedInterface = GetInterface(interfacePointer);

                    //TODO: this is bad design. The handler should be taking care of this
                    UsbInterfaces.Add(associatedInterface);

                    i++;
                }

                RegisterDefaultInterfaces();
#pragma warning restore CA2000
            }
            catch (Exception ex)
            {
                Logger.LogError(ex, Messages.ErrorMessageCouldntIntializeDevice);
                throw;
            }
        }

        private WindowsUsbInterface GetInterface(SafeFileHandle interfaceHandle)
        {
            //TODO: We need to get the read/write size from a different API call...

            //TODO: Where is the logger/tracer?
            var isSuccess = WinUsbApiCalls.WinUsb_QueryInterfaceSettings(interfaceHandle, 0, out var interfaceDescriptor);

            var retVal = new WindowsUsbInterface(interfaceHandle, interfaceDescriptor.bInterfaceNumber, Logger, ReadBufferSizeProtected, WriteBufferSizeProtected);
            WindowsDeviceBase.HandleError(isSuccess, "Couldn't query interface");

            for (byte i = 0; i < interfaceDescriptor.bNumEndpoints; i++)
            {
                isSuccess = WinUsbApiCalls.WinUsb_QueryPipe(interfaceHandle, 0, i, out var pipeInfo);
                WindowsDeviceBase.HandleError(isSuccess, "Couldn't query endpoint");
                retVal.UsbInterfaceEndpoints.Add(new WindowsUsbInterfaceEndpoint(pipeInfo.PipeId, pipeInfo.PipeType));
            }

            return retVal;
        }
        #endregion

        #region Public Methods
        public static ConnectedDeviceDefinition GetDeviceDefinition(SafeFileHandle defaultInterfaceHandle, string deviceId, ILogger logger)
        {
            var bufferLength = (uint)Marshal.SizeOf(typeof(USB_DEVICE_DESCRIPTOR));
#pragma warning disable IDE0059 // Unnecessary assignment of a value
            var isSuccess2 = WinUsbApiCalls.WinUsb_GetDescriptor(defaultInterfaceHandle, WinUsbApiCalls.DEFAULT_DESCRIPTOR_TYPE, 0, WinUsbApiCalls.EnglishLanguageID, out var _UsbDeviceDescriptor, bufferLength, out var lengthTransferred);
#pragma warning restore IDE0059 // Unnecessary assignment of a value
            WindowsDeviceBase.HandleError(isSuccess2, "Couldn't get device descriptor");

            string productName = null;
            string serialNumber = null;
            string manufacturer = null;

            if (_UsbDeviceDescriptor.iProduct > 0)
            {
                productName = WinUsbApiCalls.GetDescriptor(
                    defaultInterfaceHandle,
                    _UsbDeviceDescriptor.iProduct,
                    "Couldn't get product name",
                    logger);
            }

            if (_UsbDeviceDescriptor.iSerialNumber > 0)
            {
                serialNumber = WinUsbApiCalls.GetDescriptor(
                    defaultInterfaceHandle,
                    _UsbDeviceDescriptor.iSerialNumber,
                    "Couldn't get serial number",
                    logger);
            }

            if (_UsbDeviceDescriptor.iManufacturer > 0)
            {
                manufacturer = WinUsbApiCalls.GetDescriptor(
                    defaultInterfaceHandle,
                    _UsbDeviceDescriptor.iManufacturer,
                    "Couldn't get manufacturer",
                    logger);
            }

            return new ConnectedDeviceDefinition(
                deviceId,
                DeviceType.Usb,
                productName: productName,
                serialNumber: serialNumber,
                manufacturer: manufacturer,
                vendorId: _UsbDeviceDescriptor.idVendor,
                productId: _UsbDeviceDescriptor.idProduct,
                writeBufferSize: _UsbDeviceDescriptor.bMaxPacketSize0,
                readBufferSize: _UsbDeviceDescriptor.bMaxPacketSize0
                );
        }

        public void Close()
        {
            foreach (var usbInterface in UsbInterfaces)
            {
                usbInterface.Dispose();
            }

            UsbInterfaces.Clear();

            _DeviceHandle?.Dispose();
            _DeviceHandle = null;
        }

        public override void Dispose()
        {
            if (disposed) return;
            disposed = true;

            Close();

            base.Dispose();

            GC.SuppressFinalize(this);
        }

        public async Task InitializeAsync() => await Task.Run(Initialize);

        public Task<ConnectedDeviceDefinition> GetConnectedDeviceDefinitionAsync()
        {
            if (_DeviceHandle == null) throw new NotInitializedException();

            //TODO: Is this right?
            return Task.Run(() => DeviceBase.GetDeviceDefinitionFromWindowsDeviceId(DeviceId, DeviceType.Usb, Logger));
        }

        //TODO: make async?
        //TODO: WINUSB_SETUP_PACKET not exposed
#pragma warning disable CA1801 // Review unused parameters
#pragma warning disable IDE0060 // Remove unused parameter
        public uint SendControlOutTransfer(WINUSB_SETUP_PACKET winSetupPacket, byte[] buffer)
        {
            uint bytesWritten = 0;
            uint bufferLength = 0;
            if (buffer != null && buffer.Length > 0) bufferLength = (uint)buffer.Length;

            WinUsbApiCalls.WinUsb_ControlTransfer(_InterfaceHandle.DangerousGetHandle(), winSetupPacket, buffer, bufferLength, ref bytesWritten, IntPtr.Zero); //last pointer is overlapped structure for async operations
            return bytesWritten;
        }

        //TODO: make async?
        //TODO: WINUSB_SETUP_PACKET not exposed
        public uint SendControlInTransfer(WINUSB_SETUP_PACKET winSetupPacket)
#pragma warning restore CA1801 // Review unused parameters
#pragma warning restore IDE0060 // Remove unused parameter
        {
            uint bytesWritten = 0;
#pragma warning disable CA1825 
            var buffer = new byte[0] { };
#pragma warning restore CA1825 

            WinUsbApiCalls.WinUsb_ControlTransfer(_DeviceHandle.DangerousGetHandle(), winSetupPacket, buffer, (uint)buffer.Length, ref bytesWritten, IntPtr.Zero); //last pointer is overlapped structure for async operations
            return bytesWritten;
        }
        #endregion
    }
}