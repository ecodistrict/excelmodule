using System;
using System.Text;
using System.Collections.Generic;
using System.Threading;
using System.IO;
using System.Diagnostics;

using IMB.ByteBuffers;

namespace IMB
{
    public class TEventNameEntry
    {
        public string EventName;
        public int Publishers;
        public int Subscribers;
        public int Timers;
    }

    public class TEventEntry
    {
        public TEventEntry(TConnection aConnection, Int32 aID, string aEventName)
        {
            connection = aConnection;
            ID = aID;
            FEventName = aEventName;
            FParent = null;
            FIsPublished = false;
            FIsSubscribed = false;
            FSubscribers = false;
            FPublishers = false;
        }
        ~TEventEntry()
        {
            FStreamCache.Clear();
        }
        public enum TEventKind
        {
            ekChangeObjectEvent=0, // imb version 1
            ekStreamHeader=1,
            ekStreamBody=2,
            ekStreamTail=3,
            ekBuffer=4,
            ekNormalEvent=5,
            ekChangeObjectDataEvent=6,
            ekLogWriteLn=30,
            ekTimerCancel=40,
            ekTimerPrepare=41,
            ekTimerStart=42,
            ekTimerStop=43,
            ekTimerAcknowledgedListAdd = 45,
            ekTimerAcknowledgedListRemove = 46,
            ekTimerSetSpeed = 47,
            ekTimerTick = 48,
            ekTimerAcknowledge = 49,
            ekTimerStatusRequest = 50
        }
        public enum TLogLevel
        {
            llRemark,
            llDump,
            llNormal,
            llStart,
            llFinish,
            llPush,
            llPop,
            llStamp,
            llSummary,
            llWarning,
            llError
        }
        public const int trcInfinite = int.MaxValue;
        // private/internal
        private Int32 EventKindMask = 0x000000FF;
        private Int32 EventFlagsMask = 0x0000FF00;
        private class TStreamCacheEntry
        {
            private int FStreamID;
            private Stream FStream;
            private string FName;

            public int StreamID { get { return FStreamID; } }
            public Stream Stream { get { return FStream; } }
            public override bool Equals(Object obj)
            {
                TStreamCacheEntry SCE = obj as TStreamCacheEntry;

                if (SCE != null)
                    return FStreamID == SCE.FStreamID;
                else
                    return false;
            }
            public override int GetHashCode() { return FStreamID.GetHashCode(); }
            public TStreamCacheEntry(int aStreamID, Stream aStream, string aStreamName)
            {
                FStreamID = aStreamID;
                FStream = aStream;
                FName = aStreamName;
            }
            public string Name { get { return FName; } }
        }
        private class TStreamCache
        {
            private List<TStreamCacheEntry> FStreamCacheList = new List<TStreamCacheEntry>();

            public int Count { get { return FStreamCacheList.Count; } }
            public Stream Find(int aStreamID, out string aName)
            {
                TStreamCacheEntry SCE = new TStreamCacheEntry(aStreamID, null, null);
                int i = FStreamCacheList.IndexOf(SCE);
                if (i >= 0)
                {
                    aName = FStreamCacheList[i].Name;
                    return FStreamCacheList[i].Stream;
                }
                else
                {
                    aName = "";
                    return null;
                }
            }
            public void Clear()
            {
                FStreamCacheList.Clear();
            }
            public void Cache(int aStreamID, Stream aStream, string aStreamName)
            {
                FStreamCacheList.Add(new TStreamCacheEntry(aStreamID, aStream, aStreamName));
            }
            public void Remove(int aStreamID)
            {
                FStreamCacheList.Remove(new TStreamCacheEntry(aStreamID, null, null));
            }
        }
        private bool FIsPublished;
        private bool FIsSubscribed;
        internal string FEventName;
        internal TEventEntry FParent;
        private TStreamCache FStreamCache = new TStreamCache();
        private int TimerBasicCmd(TEventKind aEventKind, string aTimerName)
        {   
            TByteBuffer Payload = new TByteBuffer();
            Payload.Prepare(aTimerName);
            Payload.PrepareApply();
            Payload.QWrite(aTimerName);
            return SignalEvent(aEventKind, Payload.Buffer);
        }
        private int TimerAcknowledgeCmd(TEventKind aEventKind, string aTimerName, string aClientName)
        {
            TByteBuffer Payload = new TByteBuffer();
            Payload.Prepare(aTimerName);
            Payload.Prepare(aClientName);
            Payload.PrepareApply();
            Payload.QWrite(aTimerName);
            Payload.QWrite(aClientName);
            return SignalEvent(aEventKind, Payload.Buffer);
        }
        internal void Subscribe()
        {
            FIsSubscribed = true;
            // send command
            TByteBuffer Payload = new TByteBuffer();
            Payload.Prepare(ID);
            Payload.Prepare(0); // EET
            Payload.Prepare(EventName);
            Payload.PrepareApply();
            Payload.QWrite(ID);
            Payload.QWrite(0); // EET
            Payload.QWrite(EventName);
            connection.WriteCommand(TConnection.TCommands.icSubscribe, Payload.Buffer);
        }
        internal void Publish()
        {
            FIsPublished = true;
            // send command
            TByteBuffer Payload = new TByteBuffer();
            Payload.Prepare(ID);
            Payload.Prepare(0); // EET
            Payload.Prepare(EventName);
            Payload.PrepareApply();
            Payload.QWrite(ID);
            Payload.QWrite(0); // EET
            Payload.QWrite(EventName);
            connection.WriteCommand(TConnection.TCommands.icPublish, Payload.Buffer);
        }
        internal bool IsEmpty { get { return !(FIsSubscribed || FIsPublished); } }
        internal void UnSubscribe(bool aChangeLocalState = true)
        {
            if (aChangeLocalState)
                FIsSubscribed = false;
            // send command
            TByteBuffer Payload = new TByteBuffer();
            Payload.Prepare(EventName);
            Payload.PrepareApply();
            Payload.QWrite(EventName);
            connection.WriteCommand(TConnection.TCommands.icUnsubscribe, Payload.Buffer);
        }
        internal void UnPublish(bool aChangeLocalState = true)
        {
            if (aChangeLocalState)
                FIsPublished = false;
            // send command
            TByteBuffer Payload = new TByteBuffer();
            Payload.Prepare(EventName);
            Payload.PrepareApply();
            Payload.QWrite(EventName);
            connection.WriteCommand(TConnection.TCommands.icUnpublish, Payload.Buffer);
        }
        private bool FSubscribers;
        public bool Subscribers { get { return FSubscribers; } }
        private bool FPublishers;
        public bool Publishers { get { return FPublishers; } }
        internal void CopyHandlersFrom(TEventEntry aEventEntry)
        {
            OnChangeObject       = aEventEntry.OnChangeObject;
            OnFocus              = aEventEntry.OnFocus;
            OnNormalEvent        = aEventEntry.OnNormalEvent;
            OnBuffer             = aEventEntry.OnBuffer;
            OnStreamCreate       = aEventEntry.OnStreamCreate;
            OnStreamEnd          = aEventEntry.OnStreamEnd;
            OnChangeFederation   = aEventEntry.OnChangeFederation;
            OnTimerTick          = aEventEntry.OnTimerTick;
            OnTimerCmd           = aEventEntry.OnTimerCmd;
            OnChangeObjectData   = aEventEntry.OnChangeObjectData;
            OnOtherEvent         = aEventEntry.OnOtherEvent;
            OnSubAndPub          = aEventEntry.OnSubAndPub;
        }
        // dispatcher for all events
        internal void HandleEvent(TByteBuffer aPayload)
        {
            Int32 EventTick;
            Int32 EventKindInt;
            aPayload.Read(out EventTick);
            aPayload.Read(out EventKindInt);
            TEventKind eventKind = (TEventKind)(EventKindInt & EventKindMask);
            switch (eventKind)
            {
                case TEventKind.ekChangeObjectEvent:
                    HandleChangeObject(aPayload);
                    break;
                case TEventKind.ekChangeObjectDataEvent:
                    HandleChangeObjectData(aPayload);
                    break;
                case TEventKind.ekBuffer:
                    HandleBuffer(EventTick, aPayload);
                    break;
                case TEventKind.ekNormalEvent:
                    if (OnNormalEvent != null)
                        OnNormalEvent(this, aPayload);
                    break;
                case TEventKind.ekTimerTick:
                    HandleTimerTick(aPayload);
                    break;
                case TEventKind.ekTimerPrepare:
                case TEventKind.ekTimerStart:
                case TEventKind.ekTimerStop:
                    HandleTimerCmd(eventKind, aPayload);
                    break;
                case TEventKind.ekStreamHeader:
                case TEventKind.ekStreamBody:
                case TEventKind.ekStreamTail:
                    HandleStreamEvent(eventKind, aPayload);
                    break;
                default:
                    if (OnOtherEvent != null)
                        OnOtherEvent(this, EventTick, eventKind, aPayload);
                    break;
            }

        }
        // dispatchers for specific events
        private void HandleChangeObject(TByteBuffer aPayload)
        {
            if (OnFocus != null)
            {
                double X;
                double Y;
                aPayload.Read(out X);
                aPayload.Read(out Y);
                OnFocus(X, Y);
            }
            else
            {
                if (OnChangeFederation != null)
                {
                    Int32 Action;
                    Int32 NewFederationID;
                    string NewFederation;
                    aPayload.Read(out Action);
                    aPayload.Read(out NewFederationID);
                    aPayload.Read(out NewFederation);
                    OnChangeFederation(connection, NewFederationID, NewFederation);
                }
                else
                {
                    if (OnChangeObject != null)
                    {
                        Int32 Action;
                        Int32 ObjectID;
                        string Attribute;
                    
                        aPayload.Read(out Action);
                        aPayload.Read(out ObjectID);
                        aPayload.Read(out Attribute);
                        OnChangeObject(Action, ObjectID, ShortEventName, Attribute);
                    }
                }
            }
        }
        private void HandleChangeObjectData(TByteBuffer aPayload)
        {
            if (OnChangeObjectData != null)
            {
                Int32 Action;
                Int32 ObjectID;
                string Attribute;
                aPayload.Read(out Action);
                aPayload.Read(out ObjectID);
                aPayload.Read(out Attribute);
                TByteBuffer NewValues = aPayload.ReadByteBuffer();
                TByteBuffer OldValues = aPayload.ReadByteBuffer();
                OnChangeObjectData(this, Action, ObjectID, Attribute, NewValues, OldValues);
            }
        }
        private void HandleBuffer(Int32 aEventTick, TByteBuffer aPayload)
        {
            if (OnBuffer != null)
            {
                Int32 BufferID = aPayload.ReadInt32();
                TByteBuffer Buffer = aPayload.ReadByteBuffer();
                OnBuffer(this, aEventTick, BufferID, Buffer);
            }
        }
        private void HandleTimerTick(TByteBuffer aPayload)
        {
            if (OnTimerTick != null)
            {
                string TimerName;
                Int32 Tick;
                Int64 TickTime;
                Int64 StartTime;
                aPayload.Read(out TimerName);
                aPayload.Read(out Tick);
                aPayload.Read(out TickTime);
                aPayload.Read(out StartTime);
                OnTimerTick(this, TimerName, Tick, TickTime, StartTime);
            }
        }
        private void HandleTimerCmd(TEventKind aEventKind, TByteBuffer aPayload)
        {
            if (OnTimerCmd != null)
            {
                string TimerName;
                aPayload.Read(out TimerName);
                OnTimerCmd(this, aEventKind, TimerName);
            }
        }
        private void HandleStreamEvent(TEventKind aEventKind, TByteBuffer aPayload)
        {
            Int32 StreamID;
            string StreamName;
            Stream stream;
            switch (aEventKind)
            {
                case TEventKind.ekStreamHeader:
                    if (OnStreamCreate != null)
                    {
                        aPayload.Read(out StreamID);
                        aPayload.Read(out StreamName);
                        stream = OnStreamCreate(this, StreamName);
                        if (stream != null)
                            FStreamCache.Cache(StreamID, stream, StreamName);
                    }
                    break;
                case TEventKind.ekStreamBody:
                    aPayload.Read(out StreamID);
                    stream = FStreamCache.Find(StreamID, out StreamName);
                    if (stream != null)
                        stream.Write(aPayload.Buffer, aPayload.ReadCursor, aPayload.ReadAvailable);
                    break;
                case TEventKind.ekStreamTail:
                    aPayload.Read(out StreamID);
                    stream = FStreamCache.Find(StreamID, out StreamName);
                    if (stream != null)
                    {
                        stream.Write(aPayload.Buffer, aPayload.ReadCursor, aPayload.ReadAvailable);
                        if (OnStreamEnd != null)
                            OnStreamEnd(this, ref stream, StreamName);
                        stream.Close();
                        FStreamCache.Remove(StreamID);
                    }
                    break;
            }
        }
        internal void HandleOnSubAndPub(TConnection.TCommands aCommand, string aEventName, bool aIsChild)
        {
            if (!aIsChild)
            {
                switch (aCommand)
                {
                    case TConnection.TCommands.icSubscribe:
                        FSubscribers = true;
                        break;
                    case TConnection.TCommands.icPublish:
                        FPublishers = true;
                        break;
                    case TConnection.TCommands.icUnsubscribe:
                        FSubscribers = false;
                        break;
                    case TConnection.TCommands.icUnpublish:
                        FPublishers = false;
                        break;
                }
            }
            if (OnSubAndPub != null)
                OnSubAndPub(this, aCommand, aEventName, aIsChild);
        }
        // public
        public readonly TConnection connection;
        public readonly Int32 ID;
        public string EventName { get { return FEventName; } }
        public string ShortEventName
        {
            get
            {
                string federationPrefix = connection.Federation + ".";
                if (FEventName.StartsWith(federationPrefix))
                    return FEventName.Substring(federationPrefix.Length);
                else
                    return FEventName;
            }
        }
        public bool IsPublished { get { return FIsPublished; } }
        public bool IsSubscribed { get { return FIsSubscribed; } }
        // imb 1
        public delegate void TOnChangeObject(Int32 aAction, Int32 aObjectID, string aObjectName, string aAttribute);
        public delegate void TOnFocus(double x, double y);
        public event TOnChangeObject OnChangeObject = null;
        public event TOnFocus OnFocus = null;
        // imb 2
        public delegate void TOnNormalEvent(TEventEntry aEvent, TByteBuffer aPayload);
        public delegate void TOnBuffer(TEventEntry aEvent, Int32 aTick, Int32 aBufferID, TByteBuffer aBuffer);
        public delegate Stream TOnStreamCreate(TEventEntry aEvent, string aStreamName);
        public delegate void TOnStreamEnd(TEventEntry aEvent, ref Stream aStream, string aStreamName);
        public delegate void TOnChangeFederation(TConnection aConnection, Int32 aNewFederationID, string aNewFederation);
        public event TOnNormalEvent OnNormalEvent = null;
        public event TOnBuffer OnBuffer = null;
        public event TOnStreamCreate OnStreamCreate = null;
        public event TOnStreamEnd OnStreamEnd = null;
        public event TOnChangeFederation OnChangeFederation = null;
        // imb 3
        public delegate void TOnTimerTick(TEventEntry aEvent, string aTimerName, Int32 aTick, Int64 aTickTime, Int64 aStartTime);
        public delegate void TOnTimerCmd(TEventEntry aEvent, TEventKind aEventKind, string aTimerName);
        public delegate void TOnChangeObjectData(TEventEntry aEvent, Int32 aAction, Int32 aObjectID, string aAttribute, TByteBuffer aNewValues, TByteBuffer aOldValues);
        public delegate void TOnSubAndPubEvent(TEventEntry aEvent, TConnection.TCommands aCommand, string aEventName, bool aIsChild);
        public delegate void TOnOtherEvent(TEventEntry aEvent, Int32 aTick, TEventKind aEventKind, TByteBuffer aPayload);
        public event TOnTimerTick OnTimerTick = null;
        public event TOnTimerCmd OnTimerCmd = null;
        public event TOnChangeObjectData OnChangeObjectData = null;
        public event TOnSubAndPubEvent OnSubAndPub = null;
        public event TOnOtherEvent OnOtherEvent = null;
        // signals (send events)
        public int SignalEvent(TEventKind aEventKind, byte[] aEventPayload)
        {
            TByteBuffer Payload = new TByteBuffer();
            if (!IsPublished && connection.AutoPublish)
                Publish();
            if (IsPublished)
            {
                Payload.Prepare(ID);
                Payload.Prepare((Int32)0); // tick
                Payload.Prepare((Int32)aEventKind);
                Payload.Prepare(aEventPayload);
                Payload.PrepareApply();
                Payload.QWrite(ID);
                Payload.QWrite((Int32)(0)); // tick
                Payload.QWrite((Int32)aEventKind);
                Payload.QWrite(aEventPayload);
                return connection.WriteCommand(TConnection.TCommands.icEvent, Payload.Buffer);
            }
            else
                return TConnection.iceNotEventPublished;
        }
        public int SignalBuffer(Int32 aBufferID, byte[] aBuffer, Int32 aEventFlags = 0)
        {
            TByteBuffer Payload = new TByteBuffer();
            if (!IsPublished && connection.AutoPublish)
                Publish();
            if (IsPublished)
            {
                Payload.Prepare(ID);
                Payload.Prepare((Int32)0); // tick
                Payload.Prepare((Int32)TEventKind.ekBuffer | (aEventFlags & EventFlagsMask));
                Payload.Prepare(aBufferID);
                Payload.Prepare(aBuffer.Length); 
                Payload.Prepare(aBuffer);
                Payload.PrepareApply();
                Payload.QWrite(ID);
                Payload.QWrite((Int32)(0)); // tick
                Payload.QWrite((Int32)TEventKind.ekBuffer | (aEventFlags & EventFlagsMask));
                Payload.QWrite(aBufferID);
                Payload.QWrite(aBuffer.Length); 
                Payload.QWrite(aBuffer);
                return connection.WriteCommand(TConnection.TCommands.icEvent, Payload.Buffer);
            }
            else
                return TConnection.iceNotEventPublished;
        }
        private int ReadBytesFromStream(TByteBuffer aBuffer, Stream aStream)
        {
            try
            {
                int Count = 0;
                int NumBytesRead = -1;
                while (aBuffer.WriteAvailable > 0 && NumBytesRead != 0)
                {
                    NumBytesRead = aStream.Read(aBuffer.Buffer, aBuffer.WriteCursor, aBuffer.WriteAvailable);
                    aBuffer.Written(NumBytesRead);
                    Count += NumBytesRead;
                }
                return Count;
            }
            catch (IOException)
            {
                return 0; // signal stream read error
            }
        }
        
        public int SignalStream(string aStreamName, Stream aStream)
        {
            TByteBuffer Payload = new TByteBuffer();
            Int32 ReadSize;
            Int32 BodyIndex;
            Int32 EventKindIndex;
            if (!IsPublished && connection.AutoPublish)
                Publish();
            if (IsPublished)
            {
                // ekStreamHeader, includes stream name, no stream data
                byte[] StreamNameUTF8 = Encoding.UTF8.GetBytes(aStreamName);
                Int32 StreamID = connection.getConnectionHashCode(StreamNameUTF8);
                Payload.Prepare(ID);
                Payload.Prepare((Int32)0); // tick
                Payload.Prepare((Int32)TEventKind.ekStreamHeader); // event kind
                Payload.Prepare(StreamID);
                Payload.Prepare(aStreamName);
                Payload.PrepareApply();
                Payload.QWrite(ID);
                Payload.QWrite((Int32)0); // tick
                EventKindIndex = Payload.WriteCursor;
                Payload.QWrite((Int32)TEventKind.ekStreamHeader); // event kind
                Payload.QWrite(StreamID);
                BodyIndex = Payload.WriteCursor;
                Payload.QWrite(aStreamName);
                int res = connection.WriteCommand(TConnection.TCommands.icEvent, Payload.Buffer);
                if (res>0)
                {
                    // ekStreamBody, only buffer size chunks of data
                    // prepare Payload to same value but aStreamName stripped
                    // fixup event kind
                    Payload.WriteStart(EventKindIndex);
                    Payload.QWrite((Int32)TEventKind.ekStreamBody);
                    Payload.WriteStart(BodyIndex);
                    // prepare room for body data
                    Payload.PrepareStart();
                    Payload.PrepareSize(TConnection.MaxStreamBodyBuffer);
                    Payload.PrepareApply();
                    // write pointer in ByteBuffer is still at beginning of stream read buffer!
                    // but buffer is already created on correct length
                    do
                    {
                        ReadSize = ReadBytesFromStream(Payload, aStream);
                        //ReadSize = aStream.Read(Payload.Buffer, BodyIndex, Connection.MaxStreamBodyBuffer);
                        if (ReadSize == TConnection.MaxStreamBodyBuffer)
                            res = connection.WriteCommand(TConnection.TCommands.icEvent, Payload.Buffer);
                        // reset write position
                        Payload.WriteStart(BodyIndex);
                    }
                    while ((ReadSize == TConnection.MaxStreamBodyBuffer) && (res > 0));
                    if (res>0)
                    {
                        // clip ByteBuffer to bytes read from stream
                        // write pointer in ByteBuffer is still at beginning of stream read buffer!
                        Payload.PrepareStart();
                        Payload.PrepareSize(ReadSize);
                        Payload.PrepareApplyAndTrim();
                        // fixup event kind
                        Payload.WriteStart(EventKindIndex);
                        Payload.QWrite((Int32)TEventKind.ekStreamTail);
                        res = connection.WriteCommand(TConnection.TCommands.icEvent, Payload.Buffer);
                    }
                }
                return res;
            }
            else 
                return TConnection.iceNotEventPublished;

        }
        public const Int32 actionNew = 0;
        public const Int32 actionDelete = 1;
        public const Int32 actionChange = 2;
        public int SignalChangeObject(int aAction, int aObjectID, string aAttribute = "")
        {
            TByteBuffer Payload = new TByteBuffer();
            if (!IsPublished && connection.AutoPublish)
                Publish();
            if (IsPublished)
            {
                Payload.Prepare(ID);    
                Payload.Prepare((Int32)0); // tick
                Payload.Prepare((Int32)TEventKind.ekChangeObjectEvent);
                Payload.Prepare(aAction);
                Payload.Prepare(aObjectID);
                Payload.Prepare(aAttribute);
                Payload.PrepareApply();
                Payload.QWrite(ID);
                Payload.QWrite((Int32)(0)); // tick
                Payload.QWrite((Int32)TEventKind.ekChangeObjectEvent);
                Payload.QWrite(aAction);
                Payload.QWrite(aObjectID);
                Payload.QWrite(aAttribute);
                return connection.WriteCommand(TConnection.TCommands.icEvent, Payload.Buffer);
            }
            else
                return TConnection.iceNotEventPublished;
        }
        // timers
        public int TimerCreate(string aTimerName, Int64 aStartTimeUTCorRelFT, int aResolutionms, double aSpeedFactor, int aRepeatCount = trcInfinite)
        {
            TByteBuffer Payload = new TByteBuffer();
            if (!IsPublished && connection.AutoPublish)
                Publish();
            if (IsPublished)
            {
                Payload.Prepare(ID);
                Payload.Prepare(aTimerName);
                Payload.Prepare(aStartTimeUTCorRelFT);
                Payload.Prepare(aResolutionms);
                Payload.Prepare(aSpeedFactor);
                Payload.Prepare(aRepeatCount);
                Payload.PrepareApply();
                Payload.QWrite(ID);
                Payload.QWrite(aTimerName);
                Payload.QWrite(aStartTimeUTCorRelFT);
                Payload.QWrite(aResolutionms);
                Payload.QWrite(aSpeedFactor);
                Payload.QWrite(aRepeatCount);
                return connection.WriteCommand(TConnection.TCommands.icCreateTimer, Payload.Buffer);
            }
            else
                return TConnection.iceNotEventPublished;
        }
        public int TimerCancel(string aTimerName)
        {
            return TimerBasicCmd(TEventKind.ekTimerCancel, aTimerName);
        }
        public int TimerPrepare(string aTimerName)
        {
            return TimerBasicCmd(TEventKind.ekTimerPrepare, aTimerName);
        }
        public int TimerStart(string aTimerName)
        {
            return TimerBasicCmd(TEventKind.ekTimerStart, aTimerName);
        }
        public int TimerStop(string aTimerName)
        {
            return TimerBasicCmd(TEventKind.ekTimerStop, aTimerName);
        }
        public int TimerSetSpeed(string aTimerName, double aSpeedFactor)
        {
            TByteBuffer Payload = new TByteBuffer();
            Payload.Prepare(aTimerName);
            Payload.Prepare(aSpeedFactor);
            Payload.PrepareApply();
            Payload.QWrite(aTimerName);
            Payload.QWrite(aSpeedFactor);
            return SignalEvent(TEventKind.ekTimerSetSpeed, Payload.Buffer);
        }
        public int TimerAcknowledgeAdd(string aTimerName, string aClientName)
        {
            return TimerAcknowledgeCmd(TEventKind.ekTimerAcknowledgedListAdd, aTimerName, aClientName);
        }
        public int TimerAcknowledgeRemove(string aTimerName, string aClientName)
        {
            return TimerAcknowledgeCmd(TEventKind.ekTimerAcknowledgedListRemove, aTimerName, aClientName);
        }
        public int TimerAcknowledge(string aTimerName, string aClientName, int aProposedTimeStep)
        {
            TByteBuffer Payload = new TByteBuffer();
            Payload.Prepare(aClientName);
            Payload.Prepare(aTimerName);
            Payload.Prepare(aProposedTimeStep);
            Payload.PrepareApply();
            Payload.QWrite(aClientName);
            Payload.QWrite(aTimerName);
            Payload.QWrite(aProposedTimeStep);
            return SignalEvent(TEventKind.ekTimerAcknowledge, Payload.Buffer);
        }
        // log
        public int LogWriteLn(string aLine, TLogLevel aLevel)
        {
            TByteBuffer Payload = new TByteBuffer();
            if (!IsPublished && connection.AutoPublish)
                Publish();
            if (IsPublished)
            {
                Payload.Prepare((Int32)0); // client id filled in by hub
                Payload.Prepare(aLine);
                Payload.Prepare((Int32)aLevel);
                Payload.PrepareApply();
                Payload.QWrite((Int32)0); // client id filled in by hub
                Payload.QWrite(aLine);
                Payload.QWrite((Int32)aLevel);
                return SignalEvent(TEventKind.ekLogWriteLn, Payload.Buffer);
            }
            else
                return TConnection.iceNotEventPublished;
        }
        // other
        public void ClearAllStreams()
        {
            FStreamCache.Clear();
        }
    }

    public partial class TConnection
    {
        // constructors/destructor
        public TConnection(string aHost, int aPort, string aOwnerName, int aOwnerID, string aFederation = DefaultFederation, bool aIMB2Compatible = true, bool aStartReadingThread = true)
        {
            FFederation = aFederation;
            FOwnerName = aOwnerName;
            FOwnerID = aOwnerID;
            FIMB2Compatible = aIMB2Compatible;
            Open(aHost, aPort, aStartReadingThread);    
        }
        ~TConnection()
        {
            Close();
        }
        // internals/privates
        internal class TEventTranslation
        {
            public const Int32 InvalidTranslatedEventID = -1;

            private Int32[] FEventTranslation;

            public TEventTranslation()
            {
                FEventTranslation = new Int32[32];
                // mark all entries as invalid
                for (int i = 0; i < FEventTranslation.Length; i++)
                    FEventTranslation[i] = InvalidTranslatedEventID;
            }
            public Int32 TranslateEventID(Int32 aRxEventID)
            {
                if ((0 <= aRxEventID) && (aRxEventID < FEventTranslation.Length))
                    return FEventTranslation[aRxEventID];
                else
                    return InvalidTranslatedEventID;
            }
            public void SetEventTranslation(Int32 aRxEventID, Int32 aTxEventID)
            {
                if (aRxEventID >= 0)
                {
                    // grow event translation list until it can contain the requested id
                    while (aRxEventID >= FEventTranslation.Length)
                    {
                        int FormerSize = FEventTranslation.Length;
                        // resize event translation array to double the size
                        Array.Resize(ref FEventTranslation, FEventTranslation.Length * 2);
                        // mark all new entries as invalid
                        for (int i = FormerSize; i < FEventTranslation.Length; i++)
                            FEventTranslation[i] = InvalidTranslatedEventID;
                    }
                    FEventTranslation[aRxEventID] = aTxEventID;
                }
            }
            public void ResetEventTranslation(Int32 aTxEventID)
            {
                for (int i = 0; i < FEventTranslation.Length; i++)
                {
                    if (FEventTranslation[i] == aTxEventID)
                        FEventTranslation[i] = InvalidTranslatedEventID;
                }
            }
        }
        internal class TEventEntryList
        {
            public TEventEntryList(Int32 aInitialSize = 8)
            {
                FInitialSize = aInitialSize;
                FCount = 0;
            }
            private TEventEntry[] FEvents = new TEventEntry[0];
            private Int32 FInitialSize;
            private Int32 FCount = 0;
            public Int32 Count { get { return FCount; } }
            public TEventEntry GetEventEntry(Int32 aEventID)
            {
                if (0 <= aEventID && aEventID < FCount)
                    return FEvents[aEventID];
                else
                    return null;
            }
            public string GetEventName(Int32 aEventID)
            {
                if (0 <= aEventID && aEventID < FCount)
                {
                    if (FEvents[aEventID] != null)
                        return FEvents[aEventID].EventName;
                    else
                        return null;
                }
                else
                    return "";
            }
            public void SetEventName(Int32 aEventID, string aEventName)
            {
                if (0 <= aEventID && aEventID < FCount)
                {
                    if (FEvents[aEventID] != null)
                        FEvents[aEventID].FEventName = aEventName;
                }
            }
            public TEventEntry AddEvent(TConnection aConnection, string aEventName)
            {
                FCount++;
                while (FCount > FEvents.Length)
                {
                    if (FEvents.Length == 0)
                        Array.Resize(ref FEvents, FInitialSize);
                    else
                        Array.Resize(ref FEvents, FEvents.Length * 2);
                }
                FEvents[FCount - 1] = new TEventEntry(aConnection, FCount - 1, aEventName);

                return FEvents[FCount - 1];
            }
            public Int32 IndexOfEventName(string aEventName)
            {
                Int32 i = FCount - 1;
                while (i >= 0 && GetEventName(i) != aEventName)
                    i--;
                return i;
            }
            public TEventEntry EventEntryOnName(string aEventName)
            {
                Int32 i = FCount - 1;
                while (i >= 0 && GetEventName(i) != aEventName)
                    i--;
                if (i >= 0)
                    return FEvents[i];
                else
                    return null;
            }
        }
        private const string ModelStatusVarName = "ModelStatus";
        private const string msVarSepChar = "|";
        private const string EventFilterPostFix = "*";
        // consts
        internal const int MaxStreamBodyBuffer = 16 * 1024; // in bytes
        private const string FocusEventName = "Focus";
        // fields
        private string FRemoteHost = "";
        private int FRemotePort = 0;
        private Thread FThread = null;
        private TEventTranslation FEventTranslation = new TEventTranslation();
        private TEventEntryList FEventEntryList = new TEventEntryList();
        private string FFederation = DefaultFederation;
        // standard event references
        TEventEntry FFocusEvent;
        TEventEntry FChangeFederationEvent;
        TEventEntry FLogEvent;
        // time
        Int64 FBrokerAbsoluteTime;
        Int32 FBrokerTick;
        Int32 FBrokerTickDelta;
        private Int32 FUniqueClientID = 0;
        private Int32 FClientHandle = 0;
        private Int32 FOwnerID = 0;
        private string FOwnerName = "";
        static public string ConvertToHex(byte[] aBuffer)
        {
            string hex = "";
            foreach (byte b in aBuffer)
            {
                hex += string.Format("{0:x2}", (uint)System.Convert.ToUInt32(b.ToString()));
            }
            return hex;
        }
        static public string ConvertEscapes(byte[] aBuffer)
        {
            string hex = "";
            foreach (byte b in aBuffer)
            {
                if ((' ' <= b && b <= 'Z') || ('a'<=b && b<='z'))
                    hex += (char)b;
                else                
                    hex += '<'+string.Format("{0:x2}", (uint)System.Convert.ToUInt32(b.ToString()))+'>';
            }
            return hex;
        }
        private TEventEntry EventIDToEventL(Int32 aEventID)
        {
            lock (FEventEntryList)
            {
                return FEventEntryList.GetEventEntry(aEventID);
            }
        }
        private TEventEntry AddEvent(string aEventName)
        {
            TEventEntry Event;
            int EventID = 0;
            while (EventID < FEventEntryList.Count && !FEventEntryList.GetEventEntry(EventID).IsEmpty)
                EventID += 1;
            if (EventID < FEventEntryList.Count)
            {
                Event = FEventEntryList.GetEventEntry(EventID);
                Event.FEventName = aEventName;
                Event.FParent = null;
            }
            else
                Event = FEventEntryList.AddEvent(this, aEventName);
            return Event;
        }
        private TEventEntry AddEventL(string aEventName)
        {
            lock (FEventEntryList)
            {
                return AddEvent(aEventName);
            }
        }
        private TEventEntry FindOrAddEventL(string aEventName)
        {
            lock (FEventEntryList)
            {
                TEventEntry Event = FEventEntryList.EventEntryOnName(aEventName);
                if (Event == null)
                    Event = AddEvent(aEventName);
                return Event;
            }
        }
        private TEventEntry FindEventL(string aEventName)
        {
            lock (FEventEntryList)
            {
                return FEventEntryList.EventEntryOnName(aEventName);
            }
        }
        private TEventEntry FindEventParentL(string aEventName)
        {
            lock (FEventEntryList)
            {
                TEventEntry ParentEvent = null;
                int ParentEventNameLength = -1;
                string EventName;
                for (int EventID = 0; EventID < FEventEntryList.Count; EventID++)
                {
                    EventName = FEventEntryList.GetEventEntry(EventID).EventName;
                    if (EventName.EndsWith(EventFilterPostFix) && aEventName.StartsWith(EventName.Substring(0, EventName.Length-1)))
                    {
                        if (ParentEventNameLength < EventName.Length)
                        {
                            ParentEvent = FEventEntryList.GetEventEntry(EventID);
                            ParentEventNameLength = EventName.Length;
                        }
                    }
                }
                return ParentEvent;
            }
        }
        private TEventEntry FindEventAutoPublishL(string aEventName)
        {
            TEventEntry Event = FindEventL(aEventName);
            if (Event == null && AutoPublish)
                Event = Publish(aEventName, false);
            return Event;
        }
        internal int WriteCommand(TCommands aCommand, byte[] aPayload)
        {

            TByteBuffer Buffer = new TByteBuffer();

            Buffer.Prepare(MagicBytes);
            Buffer.Prepare((Int32)aCommand);
            Buffer.Prepare((Int32)0); // payload size
            if ((aPayload != null) && (aPayload.Length > 0))
            {
                Buffer.Prepare(aPayload);
                Buffer.Prepare(CheckStringMagic);
            }
            Buffer.PrepareApply();
            Buffer.QWrite(MagicBytes);
            Buffer.QWrite((Int32)aCommand);
            if ((aPayload != null) && (aPayload.Length > 0))
            {
                Buffer.QWrite((Int32)aPayload.Length);
                Buffer.QWrite(aPayload);
                Buffer.QWrite(CheckStringMagic);
            }
            else
                Buffer.QWrite((Int32)0);
            // send buffer over socket
            if (Connected)
            {
                lock (this)
                {
                    try
                    {
                        WriteCommandLow(Buffer.Buffer, Buffer.Length);
                        return Buffer.Length;
                    }
                    catch
                    {
                        // todo: remove
                        Console.WriteLine("## exception in WriteCommand");
                        Close();
                        return iceConnectionClosed;
                    }
                }
            }
            else
                return iceConnectionClosed;
        }
        private string PrefixFederation(string aName, bool aUseFederationPrefix = true)
        {
            if (FFederation.Length != 0 && aUseFederationPrefix)
                return FFederation + "." + aName;
            else
                return aName;
        }
        // command handlers
        protected void HandleCommand(TCommands aCommand, TByteBuffer aPayload)
        {
            switch (aCommand)
            {
                case TCommands.icEvent:
                    HandleCommandEvent(aPayload);
                    break;
                case TCommands.icSetVariable:
                    HandleCommandVariable(aPayload);
                    break;
                case TCommands.icSetEventIDTranslation:
                    FEventTranslation.SetEventTranslation(
                        aPayload.PeekInt32(0, TEventTranslation.InvalidTranslatedEventID),
                        aPayload.PeekInt32(sizeof(Int32), TEventTranslation.InvalidTranslatedEventID));
                    break;
                case TCommands.icUniqueClientID:
                    aPayload.Read(out FUniqueClientID);
                    aPayload.Read(out FClientHandle);
                    break;
                case TCommands.icTimeStamp:
                    // ignore for now, only when using and syncing local time (we trust hub time for now)
                    aPayload.Read(out FBrokerAbsoluteTime);
                    aPayload.Read(out FBrokerTick);
                    aPayload.Read(out FBrokerTickDelta);
                    break;
                case TCommands.icEventNames:
                    HandleEventNames(aPayload);
                    break;
                case TCommands.icEndSession:
                    Close();
                    break;
                case TCommands.icSubscribe:
                case TCommands.icPublish:
                case TCommands.icUnsubscribe:
                case TCommands.icUnpublish:
                    HandleSubAndPub(aCommand, aPayload);
                    break;    
                default:
                    HandleCommandOther(aCommand, aPayload);
                    break;
            }
        }
        private void HandleCommandEvent(TByteBuffer aPayload)
        {
            Int32 TxEventID = FEventTranslation.TranslateEventID(aPayload.ReadInt32());
            if (TxEventID != TEventTranslation.InvalidTranslatedEventID)
                EventIDToEventL(TxEventID).HandleEvent(aPayload);
            else 
                Debug.Print("## Invalid event id found in event from "+FRemoteHost);
        }
        private void HandleCommandVariable(TByteBuffer aPayload)
        {
            if (FOnVariable != null || FOnStatusUpdate != null)
            {
                string VarName = aPayload.ReadString();
                // check if it is a status update
                if (VarName.EndsWith(msVarSepChar + ModelStatusVarName, StringComparison.OrdinalIgnoreCase))
                {
                    if (FOnStatusUpdate != null)
                    {
                        VarName = VarName.Remove(VarName.Length - (msVarSepChar.Length + ModelStatusVarName.Length));
                        string ModelName = VarName.Substring(8, VarName.Length - 8);
                        string ModelUniqueClientID = VarName.Substring(0, 8);
                        aPayload.ReadInt32();
                        Int32 Status = aPayload.ReadInt32(-1);
                        Int32 Progress = aPayload.ReadInt32(-1);
                        FOnStatusUpdate(this, ModelUniqueClientID, ModelName, Progress, Status);
                    }
                }
                else
                {
                    if (FOnVariable != null)
                    {
                        TByteBuffer VarValue = aPayload.ReadByteBuffer();
                        TByteBuffer PrevValue = new TByteBuffer();
                        FOnVariable(this, VarName, VarValue.Buffer, PrevValue.Buffer);
                    }
                }
            }
        }
        private void HandleEventNames(TByteBuffer aPayload)
        {
            if (OnEventNames != null)
            {
                Int32 ec;
                aPayload.Read(out ec);
                TEventNameEntry[] EventNames = new TEventNameEntry[ec];
                for (int en = 0; en < EventNames.Length; en++)
                {
                    EventNames[en] = new TEventNameEntry();
                    EventNames[en].EventName = aPayload.ReadString();
                    EventNames[en].Publishers = aPayload.ReadInt32();
                    EventNames[en].Subscribers = aPayload.ReadInt32();
                    EventNames[en].Timers = aPayload.ReadInt32();
                }
                OnEventNames(this, EventNames);
            }
        }
        private void HandleSubAndPub(TCommands aCommand, TByteBuffer aPayload)
        {
            Int32 EventID;
            Int32 EventEntryType;

            string EventName;
            TEventEntry EE;
            bool isChild;
            switch (aCommand)
            {
                case TCommands.icSubscribe:
                case TCommands.icPublish:
                    aPayload.Read(out EventID);
                    aPayload.Read(out EventEntryType);
                    aPayload.Read(out EventName);
                    EE = FindEventL(EventName);
                    if (EE == null)
                    {
                        EE = FindEventParentL(EventName);
                        isChild = true;
                    }
                    else
                        isChild = false;
                    if (EE != null && !EE.IsEmpty)
                        EE.HandleOnSubAndPub(aCommand, EventName, isChild);
                    break;
                case TCommands.icUnsubscribe:
                case TCommands.icUnpublish:
                    aPayload.Read(out EventName);
                    EE = FindEventL(EventName);
                    if (EE == null)
                    {
                        EE = FindEventParentL(EventName);
                        isChild = true;
                    }
                    else
                        isChild = false;
                    if (EE != null && !EE.IsEmpty)
                        EE.HandleOnSubAndPub(aCommand, EventName, isChild);
                    break;
            }
        }
        private void HandleCommandOther(TCommands aCommand, TByteBuffer aPayload)
        {
            // override to implement protocol extensions
        }
        private int RequestUniqueClientID()
        {
            TByteBuffer Payload = new TByteBuffer();
            Payload.Prepare((Int32)0);
            Payload.Prepare((Int32)0);
            Payload.PrepareApply();
            Payload.QWrite((Int32)0);
            Payload.QWrite((Int32)0);
            return WriteCommand(TCommands.icUniqueClientID, Payload.Buffer);
        }
        private int SetOwner()
        {
            if (Connected)
            {
                TByteBuffer Payload = new TByteBuffer();
                Payload.Prepare(FOwnerID);
                Payload.Prepare(FOwnerName);
                Payload.PrepareApply();
                Payload.QWrite(FOwnerID);
                Payload.QWrite(FOwnerName);
                return WriteCommand(TCommands.icSetClientInfo, Payload.Buffer);
            }
            else
                return iceConnectionClosed;
        }
        public enum TConnectionState
        {
            icsUninitialized,
            icsInitialized,
            icsClient,
            icsHub,
            icsEnded,
            // room for extensions ..
            // gateway values are used over network and should be same over all connected clients/brokers
            icsGateway = 100, // equal
            icsGatewayClient = 101, // this gateway acts as a client; subscribes are not received
            icsGatewayServer = 102 // this gateway treats connected broker as client
        }
        public enum TVarPrefix
        {
            vpUniqueClientID,
            vpClientHandle
        }
        // consts
        public const string DefaultFederation = "TNOdemo";
        public const int iceConnectionClosed = -1;
        public const int iceNotEventPublished = -2;
        // fields
        public string Federation
        {
            get { return FFederation; }
            set
            {
                string OldFederation = FFederation;
                TEventEntry Event;
                if (Connected && (OldFederation.Length != 0))
                {
                    // unpublish and unsubscribe all
                    for (int i = 0; i < FEventEntryList.Count; i++)
                    {
                        string EventName = FEventEntryList.GetEventName(i);
                        if (EventName.Length != 0 && EventName.StartsWith(OldFederation + "."))
                        {
                            Event = FEventEntryList.GetEventEntry(i);
                            if (Event.IsSubscribed)
                                Event.UnSubscribe(false);
                            if (Event.IsPublished)
                                Event.UnPublish(false);
                        }
                    }
                }
                FFederation = value;
                if (Connected && (OldFederation.Length != 0))
                {
                    // publish and subscribe all
                    for (int i = 0; i < FEventEntryList.Count; i++)
                    {
                        string EventName = FEventEntryList.GetEventName(i);
                        if (EventName.Length != 0 && EventName.StartsWith(OldFederation + "."))
                        {
                            Event = FEventEntryList.GetEventEntry(i);
                            Event.FEventName = FFederation + Event.EventName.Remove(0, OldFederation.Length);
                            if (Event.IsSubscribed)
                                Event.Subscribe();
                            if (Event.IsPublished)
                                Event.Publish();
                        }
                    }
                }
            }
        }
        public bool AutoPublish = true;
        private bool FIMB2Compatible;
        // connection
        public string RemoteHost { get { return FRemoteHost; } }
        public int RemotePort { get { return FRemotePort; } }
        public bool Open(string aHost, int aPort, bool aStartReadingThread = true)
        {
            Close();
            try
            {
                FRemoteHost = aHost;
                FRemotePort = aPort;
                OpenLow(FRemoteHost, FRemotePort);
                if (aStartReadingThread && Connected)
                {
                    FThread = new Thread(ReadCommands);
                    FThread.Name = "IMB command reader";
                    FThread.Start();
                }
                if (Connected)
                {
                    if (FIMB2Compatible)
                        RequestUniqueClientID();
                    SetOwner();
                    // request all variables if delegates defined
                    if (FOnVariable != null || FOnStatusUpdate != null)
                        WriteCommand(TCommands.icAllVariables, null);
                }
                return Connected;
            }
            catch
            {
                return false;
            }
        }
        public void Close()
        {
            if (Connected)
            {
                if (OnDisconnect != null)
                    OnDisconnect(this);
                WriteCommand(TCommands.icEndSession, null);
                CloseLow();
                FThread = null;
            }
        }
        public delegate void TOnDisconnect(TConnection aConnection);
        public event TOnDisconnect OnDisconnect;
        public void SetThrottle(Int32 aThrottle)
        {
            TByteBuffer Payload = new TByteBuffer();
            Payload.Prepare(aThrottle);
            Payload.PrepareApply();
            Payload.QWrite(aThrottle);
            WriteCommand(TCommands.icSetThrottle, Payload.Buffer);
        }
        public void SetState(TConnectionState aState)
        {
            TByteBuffer Payload = new TByteBuffer();
            Payload.Prepare((Int32)aState);
            Payload.PrepareApply();
            Payload.QWrite((Int32)aState);
            WriteCommand(TCommands.icSetState, Payload.Buffer);
        }
        // owner
        public Int32 OwnerID { get { return FOwnerID; } set { if (FOwnerID != value) { FOwnerID = value; SetOwner(); } } }
        public string OwnerName { get { return FOwnerName; } set { if (FOwnerName != value) { FOwnerName = value; SetOwner(); } } }
        public Int32 UniqueClientID { get { return GetUniqueClientID(); } }
        public Int32 ClientHandle { get { return FClientHandle; } }
        // subscribe/publish
        public TEventEntry Subscribe(string aEventName, bool aUseFederationPrefix = true)
        {
            TEventEntry Event = FindOrAddEventL(PrefixFederation(aEventName, aUseFederationPrefix));
            if (!Event.IsSubscribed)
                Event.Subscribe();
            return Event;
        }
        public TEventEntry Publish(string aEventName, bool aUseFederationPrefix = true)
        {
            TEventEntry Event = FindOrAddEventL(PrefixFederation(aEventName, aUseFederationPrefix));
            if (!Event.IsPublished)
                Event.Publish();
            return Event;
        }
        public void UnSubscribe(string aEventName, bool aUseFederationPrefix = true)
        {
            TEventEntry Event = FindEventL(PrefixFederation(aEventName, aUseFederationPrefix));
            if (Event != null && Event.IsSubscribed)
                Event.UnSubscribe();
        }
        public void UnPublish(string aEventName, bool aUseFederationPrefix = true)
        {
            TEventEntry Event = FindEventL(PrefixFederation(aEventName, aUseFederationPrefix));
            if (Event != null && Event.IsPublished)
                Event.UnPublish();
        }
        public int SignalEvent(string aEventName, TEventEntry.TEventKind aEventKind, TByteBuffer aEventPayload, bool aUseFederationPrefix = true)
        {
            TEventEntry Event = FindEventAutoPublishL(PrefixFederation(aEventName, aUseFederationPrefix));
            if (Event != null)
                return Event.SignalEvent(aEventKind, aEventPayload.Buffer);
            else 
                return iceNotEventPublished;
        }
        public int SignalEvent(int aEventID, TEventEntry.TEventKind aEventKind, TByteBuffer aEventPayload, bool aUseFederationPrefix = true)
        {
              return EventIDToEventL(aEventID).SignalEvent(aEventKind, aEventPayload.Buffer);
        }
        public int SignalBuffer(string aEventName, Int32 aBufferID, byte[] aBuffer, Int32 aEventFlags = 0, bool aUseFederationPrefix = true)
        {
            TEventEntry Event = FindEventAutoPublishL(PrefixFederation(aEventName, aUseFederationPrefix));
            if (Event != null)
                return Event.SignalBuffer(aBufferID, aBuffer, aEventFlags);
            else
                return iceNotEventPublished;
        }
        public int SignalBuffer(int aEventID, Int32 aBufferID, byte[] aBuffer, Int32 aEventFlags = 0)
        {
            return EventIDToEventL(aEventID).SignalBuffer(aBufferID, aBuffer, aEventFlags);
        }
        public int SignalChangeObject(string aEventName, Int32 aAction, Int32 aObjectID, string aAttribute = "", bool aUseFederationPrefix = true)
        {
            TEventEntry Event = FindEventAutoPublishL(PrefixFederation(aEventName, aUseFederationPrefix));
            if (Event != null)
                return Event.SignalChangeObject(aAction, aObjectID, aAttribute);
            else
                return iceNotEventPublished;
        }
        public int SignalChangeObject(int aEventID, Int32 aAction, Int32 aObjectID, string aAttribute = "")
        {
            return EventIDToEventL(aEventID).SignalChangeObject(aAction, aObjectID, aAttribute);
        }
        public int SignalStream(string aEventName, string aStreamName, Stream aStream, bool aUseFederationPrefix = true)
        {
            TEventEntry Event = FindEventAutoPublishL(PrefixFederation(aEventName, aUseFederationPrefix));
            if (Event != null)
                return Event.SignalStream(aStreamName, aStream);
            else
                return iceNotEventPublished;
        }
        public int SignalStream(int aEventID, string aStreamName, Stream aStream)
        {
            return EventIDToEventL(aEventID).SignalStream(aStreamName, aStream);
        }
        // variables
        public delegate void TOnVariable(TConnection aConnection, string aVarName, byte[] aVarValue, byte[] aPrevValue);
        private event TOnVariable FOnVariable;
        public event TOnVariable OnVariable
        {
            add
            {
                FOnVariable += value;
                WriteCommand(TCommands.icAllVariables, null); // request all varibales for initial values
            }
            remove
            {
                FOnVariable -= value;
            }
        }
        public void SetVariableValue(string aVarName, string aVarValue)
        {
            TByteBuffer Payload = new TByteBuffer();
            Payload.Prepare(aVarName);
            Payload.Prepare(aVarValue);
            Payload.PrepareApply();
            Payload.QWrite(aVarName);
            Payload.QWrite(aVarValue);
            WriteCommand(TCommands.icSetVariable, Payload.Buffer);
        }
        public void SetVariableValue(string aVarName, TByteBuffer aVarValue)
        {
            TByteBuffer Payload = new TByteBuffer();
            Payload.Prepare(aVarName);
            Payload.Prepare(aVarValue);
            Payload.PrepareApply();
            Payload.QWrite(aVarName);
            Payload.QWrite(aVarValue);
            WriteCommand(TCommands.icSetVariable, Payload.Buffer);
        }
        public void SetVariableValue(string aVarName, string aVarValue, TVarPrefix aVarPrefix)
        {
            TByteBuffer Payload = new TByteBuffer();
            Payload.Prepare((Int32)aVarPrefix);
            Payload.Prepare(aVarName);
            Payload.Prepare(aVarValue);
            Payload.PrepareApply();
            Payload.QWrite((Int32)aVarPrefix);
            Payload.QWrite(aVarName);
            Payload.QWrite(aVarValue);
            WriteCommand(TCommands.icSetVariablePrefixed, Payload.Buffer);
        }
        public void SetVariableValue(string aVarName, TByteBuffer aVarValue, TVarPrefix aVarPrefix)
        {
            TByteBuffer Payload = new TByteBuffer();
            Payload.Prepare((Int32)aVarPrefix);
            Payload.Prepare(aVarName);
            Payload.Prepare(aVarValue);
            Payload.PrepareApply();
            Payload.QWrite((Int32)aVarPrefix);
            Payload.QWrite(aVarName);
            Payload.QWrite(aVarValue);
            WriteCommand(TCommands.icSetVariablePrefixed, Payload.Buffer);
        }
        public delegate void TOnStatusUpdate(TConnection aConnection, string aModelUniqueClientID, string aModelName, Int32 aProgress, Int32 aStatus);
        private event TOnStatusUpdate FOnStatusUpdate;
        public event TOnStatusUpdate OnStatusUpdate
        {
            add
            {
                FOnStatusUpdate += value;
                WriteCommand(TCommands.icAllVariables, null); // request all varibales for initial values
            }
            remove
            {
                FOnStatusUpdate -= value;
            }
        }
        private int GetUniqueClientID()
        {
            if (FUniqueClientID == 0)
            {
                RequestUniqueClientID();
                int SpinCount = 10; // 10*500 ms
                while (FUniqueClientID == 0 && SpinCount > 0)
                {
                    Thread.Sleep(500);
                    SpinCount--;
                }
                if (FUniqueClientID == 0)
                    FUniqueClientID = -1;
            }
            return FUniqueClientID;
        }
        // status for UpdateStatus
        public readonly static Int32 statusReady= 0; // R
        public readonly static Int32 statusCalculating = 1; // C
        public readonly static Int32 statusBusy = 2; // B
        public void UpdateStatus(Int32 aProgress, Int32 aStatus)
        {
            TByteBuffer Payload = new TByteBuffer();
            Payload.Prepare(aStatus);
            Payload.Prepare(aProgress);
            Payload.PrepareApply();
            Payload.QWrite(aStatus);
            Payload.QWrite(aProgress);
            if (FIMB2Compatible)
                SetVariableValue(UniqueClientID.ToString("X8") + PrefixFederation(OwnerName).ToUpper() + msVarSepChar + ModelStatusVarName, Payload);
            else
                SetVariableValue(PrefixFederation(OwnerName).ToUpper() + msVarSepChar + ModelStatusVarName, Payload, TVarPrefix.vpUniqueClientID);
        }
        public void RemoveStatus()
        {
            if (FIMB2Compatible)
                SetVariableValue(UniqueClientID.ToString("X8") + PrefixFederation(OwnerName) + msVarSepChar + ModelStatusVarName, "");
            else
                SetVariableValue(PrefixFederation(OwnerName) + msVarSepChar + ModelStatusVarName, "", TVarPrefix.vpUniqueClientID);
        }
        public event TEventEntry.TOnFocus OnFocus
        {
            add
            {
                FFocusEvent = Subscribe(FocusEventName);
                FFocusEvent.OnFocus += value;
            }
            remove
            {
                if (FFocusEvent != null)
                    FFocusEvent.OnFocus -= value;
            }
        }
        public int SignalFocus(double aX, double aY)
        {
            if (FFocusEvent == null)
                FFocusEvent = FindEventAutoPublishL(PrefixFederation(FocusEventName));
            if (FFocusEvent != null)
            {
                TByteBuffer Payload = new TByteBuffer();
                Payload.Prepare(aX);
                Payload.Prepare(aY);
                Payload.PrepareApply();
                Payload.QWrite(aX);
                Payload.QWrite(aY);
                return FFocusEvent.SignalEvent(TEventEntry.TEventKind.ekChangeObjectEvent, Payload.Buffer);
            }
            else 
                return iceNotEventPublished;
        }
        // imb 2 change federation
        private readonly string FederationChangeEventName = "META_CurrentSession";
        public event TEventEntry.TOnChangeFederation OnChangeFederation
        {
            add
            {
                FChangeFederationEvent = Subscribe(FederationChangeEventName);
                FChangeFederationEvent.OnChangeFederation += value;
            }
            remove
            {
                if (FChangeFederationEvent != null)
                    FChangeFederationEvent.OnChangeFederation -= value;
            }
        }
        public int SignalChangeFederation(Int32 aNewFederationID, string aNewFederation)
        {
            if (FChangeFederationEvent == null)
                FChangeFederationEvent = FindEventAutoPublishL(PrefixFederation(FederationChangeEventName));
            if (FChangeFederationEvent != null)
                return FChangeFederationEvent.SignalChangeObject(TEventEntry.actionChange, aNewFederationID, aNewFederation);
            else 
                return iceNotEventPublished;
        }
        // log
        public int LogWriteLn(string aLogEventName, string aLine, TEventEntry.TLogLevel aLevel)
        {
            if (FLogEvent == null)
                FLogEvent = FindEventAutoPublishL(PrefixFederation(aLogEventName));
            if (FLogEvent != null)
                return FLogEvent.LogWriteLn(aLine, aLevel);
            else 
                return iceNotEventPublished;
        }
        // remote event info
        public delegate void TOnEventnames(TConnection aConnection, TEventNameEntry[] aEventNames);
        public event TOnEventnames OnEventNames;
        public readonly Int32 efPublishers = 1;
        public readonly Int32 efSubscribers = 2;
        public readonly Int32 efTimers = 4;
        public int RequestEventname(string aEventNameFilter, Int32 aEventFilters)
        {
            TByteBuffer Payload = new TByteBuffer();
            Payload.Prepare(aEventNameFilter);
            Payload.Prepare(aEventFilters);
            Payload.PrepareApply();
            Payload.QWrite(aEventNameFilter);
            Payload.QWrite(aEventFilters);
            return WriteCommand(TCommands.icRequestEventNames, Payload.Buffer);
        }
    }
}