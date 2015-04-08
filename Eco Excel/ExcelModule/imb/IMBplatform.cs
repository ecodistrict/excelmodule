using System;
using System.Net;
using System.Net.Sockets;
using System.IO;
using System.Threading;
using System.Diagnostics;

using IMB.ByteBuffers;

namespace IMB
{
    public partial class TConnection
    {
        // shared definitions
        public static readonly byte[] MagicBytes = new byte[8] { 0x2F, 0x47, 0x61, 0x71, 0x95, 0xAD, 0xC5, 0xFB };
        protected Int64 MagicBytesInt64 = BitConverter.ToInt64(MagicBytes, 0);
        public const Int32 CheckStringMagic = (Int32)0x10F13467;

        public const int MaxPayloadSize = 10 * 1024 * 1024; // in bytes
        
        public enum TCommands
        {
            icHeartBeat = -4,
            icEndSession = -5,
            icFlushQueue = -6,
            icUniqueClientID = -7,
            icTimeStamp = -8,

            icEvent = -15,

            icEndClientSession = -21,
            icFlushClientQueue = -22,
            icConnectToGateway = -23,

            icSetClientInfo = -31,
            icSetVariable = -32,
            icAllVariables = -33,
            icSetState = -34,
            icSetThrottle = -35,
            icSetNoDelay = -36,
            icSetVariablePrefixed = -37,

            icRequestEventNames = -41,
            icEventNames = -42,
            icRequestSubscribers = -43,
            icRequestPublishers = -44,
            icSubscribe = -45,
            icUnsubscribe = -46,
            icPublish = -47,
            icUnpublish = -48,
            icSetEventIDTranslation = -49,

            icStatusEvent = -52,
            icStatusClient = -53,
            icStatusEventPlus = -54,
            icStatusClientPlus = -55,
            icStatusHUB = -56,
            icStatusTimer = -57,

            icHumanReadableHeader = -60,
            icSetMonitor = -61,
            icResetMonitor = -62,

            icCreateTimer = -73,

            // locator commands (udp)
            icHUBLocate = -81,
            icHUBFound = -82,

            icLogClear = -91,
            icLogRequest = -92,
            icLogContents = -93
        }

        // platform specific implementation
        private TcpClient FClient = new TcpClient();
        private NetworkStream FNetStream = null;

        internal Int32 getConnectionHashCode(byte[] aNameUTF8)
        {
            return unchecked(aNameUTF8.GetHashCode() + FClient.GetHashCode()+UniqueClientID);
        }
        
        public bool Connected { get { return ((FNetStream != null) && FClient.Connected); } }

        protected int ReadBytesFromNetStream(TByteBuffer aBuffer)
        {
            try
            {
                int Count = 0;
                int NumBytesRead = -1;
                while (aBuffer.WriteAvailable > 0 && NumBytesRead != 0)
                {
                    NumBytesRead = FNetStream.Read(aBuffer.Buffer, aBuffer.WriteCursor, aBuffer.WriteAvailable);
                    aBuffer.Written(NumBytesRead);
                    Count += NumBytesRead;
                }
                return Count;
            }
            catch (IOException)
            {
                return 0; // signal connection error
            }
        }

        // Debug: public Int64 ReadCommandsCount = 0;
        
        // function returns payload of command, fills found command and returns problems during read in aResult
        // commandmagic + command + payloadsize [ + payload + payloadmagic]
        protected bool ReadCommand(ref TCommands aCommand, byte[] aFixedCommandPart, TByteBuffer aPayload, byte[] aPayloadCheck)
        {
            // Debug: ReadCommandsCount++;
            int NumBytesRead = FNetStream.Read(aFixedCommandPart, 0, aFixedCommandPart.Length);
            if (NumBytesRead > 0)
            {
                while (NumBytesRead < aFixedCommandPart.Length)
                    NumBytesRead += FNetStream.Read(aFixedCommandPart, NumBytesRead, aFixedCommandPart.Length - NumBytesRead);
                while (BitConverter.ToInt64(aFixedCommandPart, 0) != MagicBytesInt64)
                {
                    Array.Copy(aFixedCommandPart, 1, aFixedCommandPart, 0, aFixedCommandPart.Length - 1);
                    int rbr = FNetStream.ReadByte();
                    if (rbr != -1)
                        aFixedCommandPart[aFixedCommandPart.Length - 1] = (byte)rbr; // skipped bytes because of invalid magic in read command
                    else
                        return false; // error, no valid connection
                }
                // we found the magic in the stream
                aCommand = (TCommands)BitConverter.ToInt32(aFixedCommandPart, MagicBytes.Length);
                Int32 PayloadSize = BitConverter.ToInt32(aFixedCommandPart, MagicBytes.Length + sizeof(Int32));
                if (PayloadSize <= MaxPayloadSize)
                {
                    aPayload.Clear(PayloadSize);
                    if (PayloadSize > 0)
                    {
                        int Len = ReadBytesFromNetStream(aPayload);
                        if (Len == aPayload.Length)
                        {
                            NumBytesRead = 0;
                            while (NumBytesRead < aPayloadCheck.Length)
                                NumBytesRead += FNetStream.Read(aPayloadCheck, NumBytesRead, aPayloadCheck.Length - NumBytesRead);
                            return (NumBytesRead == aPayloadCheck.Length) && (BitConverter.ToInt32(aPayloadCheck, 0) == CheckStringMagic);
                        }
                        else
                            return false; // error, payload size mismatch
                    }
                    else
                        return true; // ok, no payload
                }
                else
                    return false;  // error, payload is over max size
            }
            else
                return false; //  error, no valid connection
        }

        protected void ReadCommands()
        {
            // todo: more like Delphi code
            TCommands Command = TCommands.icEndSession;
            // define once
            byte[] FixedCommandPart = new byte[MagicBytes.Length + sizeof(Int32) + sizeof(Int32)]; // magic + command + payloadsize
            TByteBuffer Payload = new TByteBuffer();
            byte[] PayloadCheck = new byte[sizeof(Int32)];
            do
            {
                try
                {
                    try
                    {
                        if (ReadCommand(ref Command, FixedCommandPart, Payload, PayloadCheck))
                            HandleCommand(Command, Payload);
                        else
                            if (Connected)
                                Debug.Print("## Dropped command ("+(int)Command+") in ReadCommands loop");
                    }
                    catch (ThreadAbortException)
                    {
                        Thread.ResetAbort();
                        Command = TCommands.icEndSession;
                    }
                }
                catch (Exception e)
                {
                    if (Connected) 
                        Debug.Print("## Exception in ReadCommands loop: " + e.Message);
                }
            } while (Command != TCommands.icEndSession && Connected);
        }

        // manually reading commands when not using a reader thread
        public void ReadCommandsNonBlocking()
        {
            TCommands Command = TCommands.icEndSession;
            byte[] FixedCommandPart = new byte[MagicBytes.Length + sizeof(Int32) + sizeof(Int32)]; // magic + command + payloadsize
            TByteBuffer Payload = new TByteBuffer();
            byte[] PayloadCheck = new byte[sizeof(Int32)];
            if (FNetStream.DataAvailable)
            {
                do
                {
                    if (ReadCommand(ref Command, FixedCommandPart, Payload, PayloadCheck))
                        HandleCommand(Command, Payload);
                } while (Command != TCommands.icEndSession && Connected);
            }
        }

        // manually reading commands when not using a reader thread
        public void ReadCommandsNonThreaded(int aTimeOut)
        {
            TCommands Command = TCommands.icEndSession;
            byte[] FixedCommandPart = new byte[MagicBytes.Length + sizeof(Int32) + sizeof(Int32)]; // magic + command + payloadsize
            TByteBuffer Payload = new TByteBuffer();
            byte[] PayloadCheck = new byte[sizeof(Int32)];
            FNetStream.ReadTimeout = aTimeOut;
            do
            {
                if (ReadCommand(ref Command, FixedCommandPart, Payload, PayloadCheck))
                    HandleCommand(Command, Payload);
            } while ((Command != TCommands.icEndSession) && Connected);
        }

        protected void WriteCommandLow(byte[] aData, int aNumberOfBytes)
        {
            FNetStream.Write(aData, 0, aNumberOfBytes);
        }

        public bool NoDelay 
        { 
            get { return FClient.Client.NoDelay; } 
            set 
            {
                if (FClient.Client.NoDelay ^ value)
                {
                    FClient.Client.NoDelay = value;
                    TByteBuffer payload = new TByteBuffer();
                    payload.Write(value);
                    WriteCommand(TCommands.icSetNoDelay, payload.Buffer);
                }
            } 
        }
        
        public bool Linger { get { return FClient.Client.LingerState.Enabled; } set { FClient.Client.LingerState = new LingerOption(value, 2); } } // set linger time to 2 seconds

        protected void OpenLow(string aRemoteHost, int aRemotePort)
        {
            FClient.Connect(aRemoteHost, aRemotePort);
            FNetStream = FClient.GetStream();
        }

        protected void CloseLow()
        {
            FClient.Close();
            // cannot use old connection so create new one to make later call to OpenLow possible
            FClient = new TcpClient();
            if (FNetStream != null)
            {
                FNetStream.Close();
                FNetStream = null;
            }
        }
    }
}
