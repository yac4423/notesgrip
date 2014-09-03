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
    
    def CreateRichTextStyl()
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
    
  end
  
  # ====================================================
  # ================= NotesDbDirectory Class ===============
  # ====================================================
  class NotesDbDirectory < GripWrapper
    def nitialize(raw_doc)
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
    # ---- Additional Methods -----
    def inspect
      "<#{self.class}, #{self.LocalTime} #{self.TimeOnly}>"
    end
  end
  
  # ====================================================
  # ================= NotesLog Class ===============
  # ====================================================
  class NotesLog < GripWrapper
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

