module Notesgrip
  # ====================================================
  # ================= NotesSession Class ===============
  # ====================================================
  class NotesSession < GripWrapper
    @@ns = nil
    
    def initialize()
      if @@ns
        @raw_object = @ns
      else
        @raw_object = WIN32OLE.new('Notes.NotesSession')
      end
    end
    
    def AddressBooks()
      rawdb_arr = @raw_object.AddressBooks()
      db_arr = []
      rawdb_arr.each {|raw_db|
        db_arr.push NotesDatabase.new(raw_db)
      }
      db_arr
    end
    
    def International()
      raw_international = @raw_object.International
      NotesInternational.new(raw_international)
    end
    
    def CurrentDatabase
      db = @raw_object.CurrentDatabase
      db ? NotesDatabase.new(db) : nil
    end
    
    def SavedData
      raw_saveddb = @raw_object.SavedData
      raw_saveddb ? NotesDocument.new(raw_saveddb) : nil
    end
    
    def URLDatabase
      db = @raw_object.URLDatabase
      db ? NotesDatabase.new(db) : nil
    end
    
    def UserGroupNameList
      raw_name_arr = @raw_object.UserGroupNameList
      name_arr = []
      raw_name_arr.each {|raw_name|
        name_arr.push NotesName.new(raw_name)
      }
      name_arr
    end
    
    def UserNameList
      raw_name_arr = @raw_object.UserNameList
      name_arr = []
      raw_name_arr.each {|raw_name|
        name_arr.push NotesName.new(raw_name)
      }
      name_arr
    end
    
    def UserNameObject
      NotesName.new(@raw_object.UserNameObject)
    end
    
    def CreateAdministrationProcess( server )
      admin_process = @raw_object.CreateAdministrationProcess( server )
      NotesAdministrationProcess.new(admin_process)
    end
    
    def CreateColorObject()
      NotesColorObject.new(@raw_object.CreateColorObject())
    end
    
    def CreateDateRange()
      NotesDateRange.new(@raw_object.CreateDateRange())
    end
    
    def CreateDateTime( dateTime)
      NotesDateTime.new(@raw_object.CreateDateTime( dateTime))
    end
    
    def CreateLog(programName)
      NotesLog.new(@raw_object.CreateLog(programName))
    end
    
    def CreateName( name, language=nil )
      NotesName.new(@raw_object.CreateName(name, language))
    end
    
    def CreateNewsletter(notesDocumentCollection)
      NotesNewsletter.new(@raw_object.CreateNewsletter(toRaw(notesDocumentCollection)))
    end
    
    def CreateRegistration()
      NotesRegistration.new(@raw_object.CreateRegistration())
    end
    
    def CreateRichTextParagraphStyle()
      NotesRichTextParagraphStyle.new(@raw_object.CreateRichTextParagraphStyle())
    end
    
    def CreateRichTextStyle()
      NotesRichTextStyle.new(@raw_object.CreateRichTextStyl())
    end
    
    def CreateStream()
      NotesStream.new(@raw_object.CreateStream())
    end
    
    def Evaluate(formula, doc)
      @raw_object.Evaluate(formula, toRaw(doc))
    end
    
    def FreeTimeSearch( window, duration, names, firstfit=nil)
      dateRangeArr = @raw_object.FreeTimeSearch( window, duration, names, firstfit)
      ret_list = []
      dateRangeArr.each {|notesDateRange|
        ret_list.push NotesDateRange.new(notesDateRange)
      }
      ret_list
    end
    
    def GetDatabase(server_name, db_filename)
      raw_db = @raw_object.GetDatabase(server_name, db_filename)
      NotesDatabase.new(raw_db)
    end
    alias database GetDatabase
    
    def GetDbDirectory(server_name)
      raw_db_directory = @raw_object.GetDbDirectory(server_name)
      NotesDbDirectory.new(raw_db_directory)
    end
    
    def GetDirectory()
      NotesDirectory.new(@raw_object.GetDirectory())
    end
    
    def GetEnvironmentString( name, system=false )
      @raw_object.GetEnvironmentString( name, system)
    end
    
    def GetEnvironmentValue(name, system=false )
      @raw_object.GetEnvironmentValue(name, system)
    end
    

    POLICYSETTINGS_ARCHIVE = 2
    POLICYSETTINGS_DESKTOP = 4
    POLICYSETTINGS_REGISTRATION = 0
    POLICYSETTINGS_SECURITY = 3
    POLICYSETTINGS_SETUP = 1
    def GetUserPolicySettings( server , name , type , explicitPolicy=nil )
      doc = @raw_object.GetUserPolicySettings( server , name , type , explicitPolicy )
      NotesDocument.new(doc)
    end
    
    def HashPassword(password)
      @raw_object.HashPassword(password)
    end
    
    def InitializeUsingNotesUserName( name, password )
      @raw_object.InitializeUsingNotesUserName( name, password )
    end
    
    def Resolve( url )
      obj = @raw_object.Resolve( url )
      case url
      when /\?OpenDatabase/i
        NotesDocument.new(obj)
      when /\?OpenView/i
        NotesView.new(obj)
      when /\?OpenForm/i
        NotesForm.new(obj)
      when /\?OpenDocument/i
        doc = NotesDocument.new(obj)
      when /\?OpenAgent/i
        NotesAgent.new(obj)
      else
        obj
      end
    end
    
    def SendConsoleCommand( serverName, consoleCommand )
      @raw_object.SendConsoleCommand( serverName, consoleCommand )
    end
    
    def SetEnvironmentVar( name, valueV, issystemvar=false )
      @raw_object.SetEnvironmentVar( name, valueV, issystemvar )
    end
    
    def UpdateProcessedDoc( notesDocument )
      @raw_object.UpdateProcessedDoc( toRaw(notesDocument) )
    end
    
    def VerifyPassword( password, hashedPassword )
      @raw_object.VerifyPassword( password, hashedPassword )
    end
    
    
    # ----Additional Methods --------
    NOTES_DATABASE = 1247
    TEMPLATE = 1248
    REPLICA_CANDIDATE = 1245
    TEMPLATE_CANDIDATE = 1246 
    def each_database(serverName = "")
      db_directory = @raw_object.GetDbDirectory( serverName )
      raw_db = db_directory.GetFirstDatabase(NOTES_DATABASE)
      while raw_db
        next_db = db_directory.GetNextDatabase
        yield NotesDatabase.new(raw_db)
        raw_db = next_db
      end
    end
    
    # -------Simple Method Relay--------
    def CommonUserName()
      @raw_object.CommonUserName()
    end
    
    def ConvertMime()
      @raw_object.ConvertMime()
    end
    
    def ConvertMime=(flag)
      @raw_object.ConvertMime = flag
    end
    def CurrentAgent()
      raise "Not Support"
    end
    
    def DocumentContext()
      raise "Not Support"
    end
    
    def EffectiveUserName()
      @raw_object.EffectiveUserName()
    end
    
    def HttpURL()
      @raw_object.HttpURL()
    end
    
    def IsOnServer()
      @raw_object.IsOnServer()
    end
    
    def LastExitStatus()
      @raw_object.LastExitStatus()
    end
    
    def LastRun()
      @raw_object.LastRun()
    end
    
    def NotesBuildVersion()
      @raw_object.NotesBuildVersion()
    end
    
    def NotesURL()
      @raw_object.NotesURL()
    end
    
    def NotesVersion()
      @raw_object.NotesVersion()
    end
    
    def OrgDirectoryPath()
      @raw_object.OrgDirectoryPath()
    end
    
    def Platform()
      @raw_object.Platform()
    end
    
    def ServerName()
      @raw_object.ServerName()
    end
    
    def UserName()
      @raw_object.UserName()
    end
    
    def CreateDOMParser()
      raise "Not Support"
    end
    
    def CreateDxlExporter()
      raise "Not Support"
    end
    
    def CreateDxlImporter()
      raise "Not Support"
    end
    
    def CreateSAXParser()
      raise "Not Support"
    end
    
    def CreateTimer()
      raise "Not Support"
    end
    
    def CreateXSLTransformer()
      raise "Not Support"
    end
    
    def GetPropertyBroker()
      raise "Not Support"
    end
    
    def Initialize(password)
      @raw_object.Initialize(password)
    end
    
    def VerifyPassword(password, hashPassword)
      @raw_object.VerifyPassword(password, hashPassword)
    end
  end
  
  # ====================================================
  # ================= NotesDbDirectory Class ===============
  # ====================================================
  class NotesDbDirectory < GripWrapper
    NOTES_DATABASE = 1247
    TEMPLATE = 1248
    REPLICA_CANDIDATE = 1245
    TEMPLATE_CANDIDATE = 1246 
    def initialize(raw_doc)
      super(raw_doc)
    end
    
    def CreateDatabase(dbfile)
      servername = @raw_object.Name
      raw_db = NotesSession.new.GetDatabase(servername, dbfile)
      unless raw_db.IsOpen
        raw_db.Create(servername, dbfile, true)
      end
      NotesDatabase.new(raw_db)
    end
    
    def GetFirstDatabase(fileType=NOTES_DATABASE)
      raw_db = @raw_object.GetFirstDatabase(fileType)
      NotesDatabase.new(raw_db)
    end
    
    
    def GetNextDatabase()
      raw_db = @raw_object.GetNextDatabase()
      return nil unless raw_db
      NotesDatabase.new(raw_db)
    end
    
    def OpenDatabase( dbfile, failover=false )
      raw_db = @raw_object.OpenDatabase( dbfile, failover )
      NotesDatabase.new(raw_db)
    end
    
    def OpenDatabaseByReplicaID( repid )
      raw_db = @raw_object.OpenDatabaseByReplicaID( repid )
      NotesDatabase.new(raw_db)
    end
    
    def OpenDatabaseIfModified( dbfile , notesDateTime )
      raw_db = @raw_object.OpenDatabaseIfModified( dbfile , notesDateTime )
      NotesDatabase.new(raw_db)
    end
    
    def OpenMailDatabase()
      raw_db = @raw_object.OpenMailDatabase()
      NotesDatabase.new(raw_db)
    end
    
    # ----Additional Methods --------
    def each_database(fileType=NOTES_DATABASE)
      db = GetFirstDatabase(fileType)
      while db
        yield db
        db = GetNextDatabase()
      end
    end
  end
  
  # ====================================================
  # ================= NotesDateTime Class ===============
  # ====================================================
  class NotesDateTime < GripWrapper
    # -------Simple Method Relay--------
    def DateOnly()
      @raw_object.DateOnly()
    end
    
    def GMTTime()
      @raw_object.GMTTime()
    end
    
    def IsDST()
      @raw_object.IsDST()
    end
    
    def IsValidDate()
      @raw_object.IsValidDate()
    end
    
    def LocalTime()
      @raw_object.LocalTime()
    end
    
    def LSGMTTime()
      @raw_object.LSGMTTime()
    end
    
    def LSLocalTime()
      @raw_object.LSLocalTime()
    end
    
    def Parent()
      NotesSession.new
    end
    
    def TimeOnly()
      @raw_object.TimeOnly()
    end
    
    def TimeZone()
      @raw_object.TimeZone()
    end
    
    def ZoneTime()
      @raw_object.ZoneTime()
    end
    
    def AdjustDay()
      @raw_object.AdjustDay()
    end
    
    def AdjustHour()
      @raw_object.AdjustHour()
    end
    
    def AdjustMinute()
      @raw_object.AdjustMinute()
    end
    
    def AdjustMonth()
      @raw_object.AdjustMonth()
    end
    
    def AdjustSecond()
      @raw_object.AdjustSecond()
    end
    
    def AdjustYear()
      @raw_object.AdjustYear()
    end
    
    def ConvertToZone()
      @raw_object.ConvertToZone()
    end
    
    def SetAnyDate()
      @raw_object.SetAnyDate()
    end
    
    def SetAnyTime()
      @raw_object.SetAnyTime()
    end
    
    def SetNow()
      @raw_object.SetNow()
    end
    
    def TimeDifference(dateTime)
      raw_datetime = toRwa(dateTime)
      @raw_object.TimeDifference(raw_datetime)
    end
    
    def TimeDifferenceDouble(dateTime)
      raw_datetime = toRwa(dateTime)
      @raw_object.TimeDifferenceDouble(raw_datetime)
    end
    
    # ---- Additional Methods -----
    def inspect
      "<#{self.class}, #{self.LocalTime} #{self.TimeOnly}>"
    end
  end
  
  # ====================================================
  # ================= NotesLog Class ===============
  # ====================================================
  class NotesLog < GripWrapper
    # -------Simple Method Relay--------
    def LogActions()
      @raw_object.LogActions()
    end
    def LogActions=(flag)
      @raw_object.LogActions = flag
    end
    
    def LogErrors()
      @raw_object.LogErrors()
    end
    def LogErrors=(flag)
      @raw_object.LogErrors = flag
    end
    
    def NumActions()
      @raw_object.NumActions()
    end
    
    def NumErrors()
      @raw_object.NumErrors()
    end
    
    def OverwriteFile()
      @raw_object.OverwriteFile()
    end
    def OverwriteFile=(flag)
      @raw_object.OverwriteFile = flag
    end
    
    def Parent()
      NotesSession.new
    end
    
    def ProgramName()
      @raw_object.ProgramName()
    end
    def ProgramName=(prog_name)
      @raw_object.ProgramName = prog_name
    end
    
    def Close()
      @raw_object.Close()
    end
    
    def LogAction(description)
      @raw_object.LogAction(description)
    end
    
    def LogError(code, description)
      @raw_object.LogError(code, description)
    end
    
    def LogEvent()
      @raw_object.LogEvent()
    end
    
    def New()
      @raw_object.New()
    end
    
    def OpenAgentLog()
      @raw_object.OpenAgentLog()
    end
    
    def OpenFileLog()
      @raw_object.OpenFileLog()
    end
    
    def OpenMailLog()
      @raw_object.OpenMailLog()
    end
    
    def OpenNotesLog()
      @raw_object.OpenNotesLog()
    end
  end
  
  # ====================================================
  # ================= NotesName Class ===============
  # ====================================================
  class NotesName < GripWrapper
    def inspect
      "<#{self.class}, #{self.Common}>"
    end
  end
  
  # ====================================================
  # ================= NotesStream Class ===============
  # ====================================================
  class NotesStream < GripWrapper
  end
  
  # ====================================================
  # ================= NotesColorObject Class ===============
  # ====================================================
  class NotesColorObject < GripWrapper
  end
  
  # ====================================================
  # ======== NotesAdministrationProcess Class ===============
  # ====================================================
  class NotesAdministrationProcess < GripWrapper
  end
  
  # ====================================================
  # ======== NotesInternational Class ===============
  # ====================================================
  class NotesInternational < GripWrapper
  end
  
  # ====================================================
  # ======== NotesNewsLetter Class ===============
  # ====================================================
  class NotesNewsLetter < GripWrapper
  end
  
  # ====================================================
  # ======== NotesRegistration Class ===============
  # ====================================================
  class NotesRegistration < GripWrapper
  end
end

