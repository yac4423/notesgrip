module Notesgrip
  # ====================================================
  # ================= NotesDatabase Class ===============
  # ====================================================
  class NotesDatabase < GripWrapper
    def initialize(raw_object)
      super(raw_object)
      # auto open Database 
      unless @raw_object.isOpen
        begin
          @raw_object.Open("","")
        rescue WIN32OLERuntimeError
          # unable access DataBase
        end
      end
    end
    
    def ACL
      NotesACL.new(@raw_object.ACL)
    end
    
    def Agents
      agent_arr = @raw_object.Agents
      return [] unless agent_arr  # Case of No Agent in database
      ret_list = []
      agent_arr.each {|raw_agent|
        ret_list.push NotesAgent.new(raw_agent)
      }
      ret_list
    end
    
    def AllDocuments
      NotesDocumentCollection.new(@raw_object.AllDocuments)
    end
    
    ACLLEVEL_NOACCESS = 0
    ACLLEVEL_DEPOSITOR = 1
    ACLLEVEL_READER = 2
    ACLLEVEL_AUTHOR = 3
    ACLLEVEL_EDITOR = 4
    ACLLEVEL_DESIGNER = 5
    ACLLEVEL_MANAGER = 6
    
    def Forms
      ret_list = []
      @raw_object.Forms.each {|raw_form|
        ret_list.push NotesForm.new(raw_form)
      }
      ret_list
    end
    
    def Parent()
      NotesSession.new()
    end
    
    def ReplicationInfo
      NotesReplication.new(@raw_object.NotesReplication)
    end
    
    def Views
      ret_list = []
      raw_views = @raw_object.Views
      return [] unless raw_views
      raw_views.each {|raw_view|
        ret_list.push NotesView.new(raw_view)
      }
      ret_list
    end
    
    def CreateCopy( newServer, newDbFile , maxsize=4 )
      db = @raw_object.CreateCopy( newServer, newDbFile , maxsize )
      NotesDatabase.new(db)
    end
    
    def CreateDocument(formName=nil)
      raw_doc = @raw_object.CreateDocument()
      new_doc = raw_doc ? NotesDocument.new(raw_doc) : nil
      if new_doc and formName
        new_doc.AppendItemValue('Form', formName)
        #new_doc['Form'].Values =  formName
      end
      new_doc
    end
    
    def CreateFromTemplate( newServer, newDbFile, inheritFlag, maxsize=4 )
      db = @raw_object.CreateFromTemplate( newServer, newDbFile, inheritFlag, maxsize )
      NotesDatabase.new(db)
    end
    
    
    def CreateNoteCollection( selectAllFlag = false)
      raw_notesNotesCollection = @raw_object.CreateNoteCollection( selectAllFlag)
      NotesNoteCollection.new(raw_notesNotesCollection)
    end
    
    def CreateOutline( outlinename, defaultOutline=false )
      raw_outline = @raw_object.CreateOutline( outlinename, defaultOutline)
      NotesOutline.new(raw_outline)
    end
    
    def CreateReplica( newServer, newDbFile )
      raw_db = @raw_object.CreateReplica( newServer, newDbFile )
      raw_doc ? NotesDocument.new(raw_doc) : nil
    end
    
    def CreateView( viewName=nil, viewSelectionFormula=nil, templateView=nil, prohibitDesignRefreshModifications=false )
      raw_view = @raw_object.CreateView( viewName, viewSelectionFormula, templateView, prohibitDesignRefreshModifications )
      NotesView.new(raw_view)
    end
    
    def FTSearch( query, maxdocs=0, sortoptions=FT_SCORES, otheroptions=0 )
      raw_docCollection = @raw_object.FTSearch( query, maxdocs, sortoptions, otheroptions)
      NotesDocumentCollection.new(raw_docCollection)
    end
    
    def FTSearchRange( query, maxdocs=0, sortoptions=FT_SCORES, otheroptions=0, start=0 )
      raw_docCollection = @raw_object.FTSearchRange( query, maxdocs, sortoptions, otheroptions, start)
      NotesDocumentCollection.new(raw_docCollection)
    end
    
    def GetAgent( agentName )
      raw_agent = @raw_object.GetAgent(agentName)
      NotesAgent.new(raw_agent)
    end
    
    def GetAllReadDocuments( username=nil )
      raw_docCollection = @raw_object.GetAllReadDocuments(username)
      NotesDocumentCollection.new(raw_docCollection)
    end
    
    def GetAllUnreadDocuments( username=nil )
      raw_docCollection = @raw_object.GetAllUnreadDocuments(username)
      NotesDocumentCollection.new(raw_docCollection)
    end
    
    def GetDocumentByID( noteID )
      raw_doc = @raw_object.GetDocumentByID( noteID )
      raw_doc ? NotesDocument.new(raw_doc) : nil
    end
    
    def GetDocumentByUNID(unid)
      raw_doc = @raw_object.GetDocumentByUNID(unid)
      raw_doc ? NotesDocument.new(raw_doc) : nil
    end
    
    def GetDocumentByURL( url, reload=0, urllist=0, charset="", webusername="", webpassword=nil, proxywebusername=nil, proxywebpassword=nil, returnimmediately=false )
      #raw_doc = @raw_object.GetDocumentByURL( url, reload, urllist, charset, webusername, webpassword, proxywebusername, proxywebpassword, returnimmediately )
      #raw_doc ? NotesDocument.new(raw_doc) : nil
      raise "GetDocumentByURL() is not Work"
    end
    
    def GetForm( name )
      raw_form = @raw_object.GetForm(name)
      raw_form ? NotesForm.new(raw_form) : nil
    end
    

    DBMOD_DOC_ACL = 64
    DBMOD_DOC_AGENT = 512
    DBMOD_DOC_ALL = 32767
    DBMOD_DOC_DATA = 1
    DBMOD_DOC_FORM = 4
    DBMOD_DOC_HELP = 256
    DBMOD_DOC_ICON = 16
    DBMOD_DOC_REPLFORMULA = 2048
    DBMOD_DOC_SHAREDFIELD = 1024
    DBMOD_DOC_VIEW = 8
    def GetModifiedDocuments( since=nil , noteClass=DBMOD_DOC_DATA )
      #raw_docCollection = @raw_object.GetModifiedDocuments( since, noteClass)
      #NotesDocumentCollection.new(raw_docCollection)
      raise "GetModifiedDocuments() is not work."
    end
    
    
    def GetOutline( outlinename )
      #raw_outline = @raw_object.GetOutline( outlinename )
      #NotesOutline.new(raw_outline)
      raise "GetOutline is not work."
    end
    
    def GetProfileDocCollection( profilename=nil )
      raw_docCollection = @raw_object.GetProfileDocCollection( profilename )
      NotesDocumentCollection.new(raw_docCollection)
    end
    
    def GetProfileDocument( profilename, uniqueKey=nil )
      raw_document = @raw_object.GetProfileDocument( profilename, uniqueKey)
      raw_document ? NotesDocument.new(raw_document) : nil
    end
    
    def GetView( viewName )
      raw_view = @raw_object.GetView( viewName )
      raw_view ? NotesView.new(raw_view) : nil
    end
    alias view GetView
    
    def Search( formula, notesDateTime=nil, maxDocs=0 )
      #raw_docCollection = @raw_object.Search( formula, notesDateTime, maxDocs )
      #NotesDocumentCollection.new(raw_docCollection)
      raise "Search() is not work."
    end
    
    # ---- Additional Methods ------
    def name
      @raw_object.Title
    end
    
    def open?()
      @raw_object.IsOpen()
    end
    
    def each_document
      doc_collection = self.AllDocuments
      doc_collection.each {|doc|
        yield doc
      }
    end
    
    def each_view
      view_collection = self.Views
      view_collection.each {|view|
        yield view
      }
    end
    
    def each_form
      form_collection = self.Forms
      form_collection.each {|form|
        yield form
      }
    end
    
    def each_agent
      self.Agents.each {|agent|
        yield agent
      }
    end
    
    def each_profile
      self.GetProfileDocCollection.each {|profileDoc|
        yield profileDoc
      }
    end
    
    def inspect
      "<#{self.class}, Name:#{self.name.inspect}, FilePath:#{self.FilePath.inspect}>"
    end
    
    # -------Simple Method Relay--------
    def ACLActivityLog
      @raw_object.ACLActivityLog()
    end
    
    def Categories()
      @raw_object.Categories()
    end
    
    def Categories=(categoryList)
      @raw_object.Categories = categoryList
    end
    
    def Created()
      @raw_object.Created()
    end
    
    def CurrentAccessLevel()
      @raw_object.CurrentAccessLevel()
    end
    
    def DelayUpdates()
      @raw_object.DelayUpdates()
    end
    
    def DelayUpdates=(flag)
      @raw_object.DelayUpdates = flag
    end
    
    def DesignTemplateName()
      @raw_object.DesignTemplateName()
    end
    
    def FileFormat()
      @raw_object.FileFormat()
    end
    
    def FileName()
      @raw_object.FileName()
    end
    
    def FilePath()
      @raw_object.FilePath()
    end
    
    def FolderReferencesEnabled()
      @raw_object.FolderReferencesEnabled()
    end
    
    def FolderReferencesEnabled=(flag)
      @raw_object.FolderReferencesEnabled = flag
    end
    
    def FTIndexFrequency()
      @raw_object.FTIndexFrequency()
    end
    
    FTINDEX_DAILY = 1
    FTINDEX_HOURLY = 3
    FTINDEX_IMMEDIATE = 4
    FTINDEX_SCHEDULED = 2
    def FTIndexFrequency=(frequency)
      @raw_object.FTIndexFrequency = frequency
    end
    
    def HttpURL()
      @raw_object.HttpURL()
    end
    
    def IsClusterReplication()
      @raw_object.IsClusterReplication()
    end
    
    def IsClusterReplication=(flag)
      @raw_object.IsClusterReplication = flag
    end
    
    def IsConfigurationDirectory()
      @raw_object.IsConfigurationDirectory()
    end
    
    def IsCurrentAccessPublicReader()
      @raw_object.IsCurrentAccessPublicReader()
    end
    
    def IsCurrentAccessPublicWriter()
      @raw_object.IsCurrentAccessPublicWriter()
    end
    
    def IsDesignLockingEnabled()
      @raw_object.IsDesignLockingEnabled()
    end
    
    def IsDesignLockingEnabled=(flag)
      @raw_object.IsDesignLockingEnabled = flag
    end
    
    def IsDirectoryCatalog()
      @raw_object.IsDirectoryCatalog()
    end
    
    def IsDocumentLockingEnabled()
      @raw_object.IsDocumentLockingEnabled()
    end
    
    def IsDocumentLockingEnabled=(flag)
      @raw_object.IsDocumentLockingEnabled = flag
    end
    
    def IsFTIndexed()
      @raw_object.IsFTIndexed()
    end
    
    def IsInMultiDbIndexing()
      @raw_object.IsInMultiDbIndexing()
    end
    
    def IsInMultiDbIndexing=(flag)
      @raw_object.IsInMultiDbIndexing = flag
    end
    
    def IsInService()
      @raw_object.IsInService()
    end
    
    def IsInService=(flag)
      @raw_object.IsInService = flag
    end
    
    def IsLink()
      @raw_object.IsLink()
    end
    
    def IsMultiDbSearch()
      @raw_object.IsMultiDbSearch()
    end
    
    def IsOpen()
      @raw_object.IsOpen()
    end
    
    def IsPendingDelete()
      @raw_object.IsPendingDelete()
    end
    
    def IsPrivateAddressBook()
      @raw_object.IsPrivateAddressBook()
    end
    
    def IsPublicAddressBook()
      @raw_object.IsPublicAddressBook()
    end
    
    def LastFixup()
      @raw_object.LastFixup()
    end
    
    def LastFTIndexed()
      @raw_object.LastFTIndexed()
    end
    
    def LastModified()
      @raw_object.LastModified()
    end
    
    def LimitRevisions()
      @raw_object.LimitRevisions()
    end
    
    def LimitRevisions=(revisions)
      @raw_object.LimitRevisions = revisions
    end
    
    def LimitUpdatedBy()
      @raw_object.LimitUpdatedBy()
    end
    
    def LimitUpdatedBy=(revisions)
      @raw_object.LimitUpdatedBy = revisions
    end
    
    def ListInDbCatalog()
      @raw_object.ListInDbCatalog()
    end
    
    def ListInDbCatalog=(flag)
      @raw_object.ListInDbCatalog = flag
    end
    
    def Managers()
      @raw_object.Managers()
    end
    
    def MaxSize()
      @raw_object.MaxSize()
    end
    
    def NotesURL()
      @raw_object.NotesURL()
    end
    
    def PercentUsed()
      @raw_object.PercentUsed()
    end
    
    def ReplicaID()
      @raw_object.ReplicaID()
    end
    
    def Server()
      @raw_object.Server()
    end
    
    def Size()
      @raw_object.Size()
    end
    
    def SizeQuota()
      @raw_object.SizeQuota()
    end
    
    def SizeQuota=(quota)
      @raw_object.SizeQuota = quota
    end
    
    def SizeWarning()
      @raw_object.SizeWarning()
    end
    
    def SizeWarning=(warning)
      @raw_object.SizeWarning = warning
    end
    
    def TemplateName()
      @raw_object.TemplateName()
    end
    
    def Title()
      @raw_object.Title()
    end
    
    DBTYPE_ADDR_BOOK = 10      # Domino Directory or Personal Address Book
    DBTYPE_IMAP_SVR_PROXY = 6  # IMAP server proxy
    DBTYPE_LIBRARY = 12        # database library
    DBTYPE_LIGHT_ADDR_BOOK = 9 # Directory Catalog
    DBTYPE_MAILBOX = 3         # mailbox
    DBTYPE_MAILFILE = 2        # mail file
    DBTYPE_MULTIDB_SRCH = 8    # Domain Catalog
    DBTYPE_NEWS_SVR_PROXY = 5  # news server proxy
    DBTYPE_PERS_JOURNAL = 11   # Personal Journal
    DBTYPE_PORTFOLIO = 7       # portfolio
    DBTYPE_STANDARD = 13       # standard
    DBTYPE_SUBSCRIPTIONS = 4   # subscriptions
    DBTYPE_WEB_APP = 1         # Web application
    def Type()
      @raw_object.Type()
    end
    
    def UndeleteExpireTime()
      @raw_object.UndeleteExpireTime()
    end
    
    def UndeleteExpireTime=(hours)
      @raw_object.UndeleteExpireTime = hours
    end
    
    def UnprocessedDocuments()
      #@raw_object.UnprocessedDocuments()
      raise "UnprocessedDocuments() is not work."
    end
    
    def Compact()
      sizeDelta = @raw_object.Compact()
      return sizeDelta
    end
    
    CMPC_ARCHIVE_DELETE_COMPACT = 1          # archive and delete, then compact
    CMPC_ARCHIVE_DELETE_ONLY = 2             # archive and delete with no compact; supersedes a
    CMPC_CHK_OVERLAP = 32768                 # check overlap
    CMPC_COPYSTYLE = 16                      # copy style; supersedes b and B
    CMPC_DISABLE_DOCTBLBIT_OPTMZN = 128      # disable document table bit map optimization
    CMPC_DISABLE_LARGE_UNKTBL = 4096         # disable large unknown table
    CMPC_DISABLE_RESPONSE_INFO = 512         # disable "Don't support specialized response hierarchy"
    CMPC_DISABLE_TRANSACTIONLOGGING = 262144 # disable transaction logging
    CMPC_DISABLE_UNREAD_MARKS = 1048576      # disable "Don't maintain unread marks"
    CMPC_DISCARD_VIEW_INDICES = 32           # discard view indexes
    CMPC_ENABLE_DOCTBLBIT_OPTMZN = 64        # enable document table bit map optimization; supersedes f
    CMPC_ENABLE_LARGE_UNKTBL = 2048          # enable large unknown table; supersedes k
    CMPC_ENABLE_RESPONSE_INFO = 256          # enable "Don't support specialized response hierarchy"; supersedes H
    CMPC_ENABLE_TRANSACTIONLOGGING = 131072  # enable transaction logging; supersedes t
    CMPC_ENABLE_UNREAD_MARKS = 524288        # enable "Don't maintain unread marks"; supersedes U
    CMPC_IGNORE_COPYSTYLE_ERRORS = 1024      # ignore copy-style errors
    CMPC_MAX_4GB = 16384                     # set maximum database size at 4 gigabytes
    CMPC_NO_LOCKOUT = 8192                   # do not lock out users
    CMPC_RECOVER_INPLACE = 8                 # recover unused space in-place and reduce file size; supersedes b
    CMPC_RECOVER_REDUCE_INPLACE = 4          # recover unused space in-place without reducing file size
    CMPC_REVERT_FILEFORMAT = 65536           # do not convert old file format
    def CompactWithOptions( options )
      @raw_object.CompactWithOptions( options )
    end
    
    def Create()
      #@raw_object.Create()
      raise "Create() is not work."
    end
    
    FTINDEX_ALL_BREAKS = 4          # index sentence and paragraph breaks
    FTINDEX_ATTACHED_BIN_FILES = 16 # index attached files (binary)
    FTINDEX_ATTACHED_FILES = 1      # index attached files (raw text)
    FTINDEX_CASE_SENSITIVE = 8      # enable case-sensitive searches
    FTINDEX_ENCRYPTED_FIELDS = 2    # index encrypted fields
    def CreateFTIndex( options , recreate=false )
      @raw_object.CreateFTIndex( options , recreate )
    end
    
    def EnableFolder(foldername)
      @raw_object.EnableFolder(foldername)
    end
    
    FIXUP_INCREMENTAL = 4 # checks only documents since last Fixup
    FIXUP_NODELETE = 16   # prevents Fixup from deleting corrupted documents
    FIXUP_NOVIEWS = 64    # does not check views
    FIXUP_QUICK = 2       # checks documents more quickly but less thoroughly
    FIXUP_REVERT = 32     # reverts ID tables to the previous release format
    FIXUP_TXLOGGED = 8    # includes databases enabled for transaction logging
    FIXUP_VERIFY = 1      # makes no modifications
    def Fixup(options)
      @raw_object.Fixup(options)
    end
    
    # Constant for FTSearch
    FT_DATE_ASC = 64     # sorts by document creation date in ascending order.
    FT_DATE_DES = 32     # sorts by document creation date in descending order.
    FT_SCORES = 8        # sorts by relevance score (default).
    FT_DATABASE = 8192   # includes Domino databases.
    FT_FUZZY = 16384     # searches for related words. Need not be an exact match.
    FT_FILESYSTEM = 4096 # includes files that are not Domino databases.
    FT_STEMS = 512       # uses stem words as the basis of the search.
    FT_THESAURUS = 1024  # uses thesaurus synonyms.
    def FTDomainSearch( query, maxDocs=0, sortoptions=FT_SCORES, otheroptions=0, start=0, count=1, entryform="" )
      #raw_doc = @raw_object.FTDomainSearch( query, maxDocs, sortoptions, otheroptions, start, count, entryform )
      #raw_doc ? NotesDocument.new(raw_doc) : nil
      raise "FTDomainSearch() is not work."
    end
    
    DBOPT_LZCOMPRESSION = 65                 # uses LZ1 compression for attachments
    DBOPT_MAINTAINLASTACCESSED = 44          # maintains LastAccessed property
    DBOPT_MOREFIELDS = 54                    # allows more fields in database
    DBOPT_NOHEADLINEMONITORS = 46            # doesn't allow headline monitoring
    DBOPT_NOOVERWRITE = 36                   # doesn't overwrite free space
    DBOPT_NORESPONSEINFO = 38                # doesn't support specialized response hierarchy
    DBOPT_NOTRANSACTIONLOGGING = 45          # disables transaction logging
    DBOPT_NOUNREAD = 37                      # doesn't maintain unread marks
    DBOPT_OPTIMIZATION = 41                  # enables document table bitmap optimization
    DBOPT_REPLICATEUNREADMARKSTOANY = 71     # replicates unread marks to all servers
    DBOPT_REPLICATEUNREADMARKSTOCLUSTER = 70 # replicates unread marks to clustered servers only
    DBOPT_SOFTDELETE = 49                    # allows soft deletions
    def GetOption(optionName)
      @raw_object.GetOption(optionName)
    end
    
    def GrantAccess(name, level)
      @raw_object.GrantAccess(name, level)
    end
    
    def MarkForDelete()
      @raw_object.MarkForDelete()
    end
    
    def Open()
      @raw_object.Open()
    end
    
    def OpenByReplicaID()
      raise "OpenByReplicaID() is not work."
    end
    
    def OpenIfModified()
      raise "OpenIfModified() is not work."
    end
    
    def OpenMail()
      raise "OpenIfModified() is not work."
    end
    
    def OpenURLDb()
      raise "OpenURLDb() is not work."
    end
    
    def OpenWithFailover()
      raise "OpenWithFailover() is not work."
    end
    
    def QueryAccess(name)
      @raw_object.QueryAccess(name)
    end
    
    def QueryAccessPrivileges(name)
      @raw_object.QueryAccessPrivileges(name)
    end
    
    def QueryAccessRoles(name)
      @raw_object.QueryAccessRoles(name)
    end
    
    def Remove()
      @raw_object.Remove()
    end
    
    def RemoveFTIndex()
      @raw_object.RemoveFTIndex()
    end
    
    def Replicate(serverName)
      @raw_object.Replicate(serverName)
    end
    
    def RevokeAccess(name)
      @raw_object.RevokeAccess(name)
    end
    
    def SetOption( optionName, flag )
      @raw_object.SetOption(optionName, flag)
    end
    
    DBSIGN_DOC_ACL = 64    # signs the ACL
    DBSIGN_DOC_AGENT = 512 # signs all agents
    DBSIGN_DOC_ALL = 32767 # (default) signs all elements except data documents' active content
    DBSIGN_DOC_DATA = 1    # signs all data documents' active content (hotspots)
    DBSIGN_DOC_FORM = 4    # signs all forms
    DBSIGN_DOC_HELP = 256  # signs the "About Database" and "Using Database" documents
    DBSIGN_DOC_ICON = 16   # signs the icon
    DBSIGN_DOC_REPLFORMULA = 2048 # signs the replication formula
    DBSIGN_DOC_SHAREDFIELD = 1024 # signs all shared fields
    DBSIGN_DOC_VIEW = 8           # signs all views
    def Sign(documentType)
      @raw_object.Sign(documentType)
    end
    
    def UnprocessedFTSearch(query, maxdocs=0, sortoptions=nil, otheroptions=nil )
      raise "UnprocessedFTSearch() is not work."
    end
    
    def UnprocessedFTSearchRange(query, maxdocs=0, sortoptions=nil, otheroptions=nil, start=0)
      raise "UnprocessedFTSearchRange() is not work."
    end
    
    def UnprocessedSearch(formula, notesDateTime, maxDocs)
      raise "UnprocessedSearch() is not work."
    end
    
    def UpdateFTIndex(createFlag)
      @raw_object.UpdateFTIndex(createFlag)
    end
    
  end
  
  # ====================================================
  # ================= NotesForm Class ===============
  # ====================================================
  class NotesForm < GripWrapper
    # ------ Additional Methods -----
    def inspect
      "<#{self.class} Name:#{self.Name.inspect}>"
    end
  end
  
  # ====================================================
  # ================= NotesAgent Class ===============
  # ====================================================
  class NotesAgent < GripWrapper
    # ------ Additional Methods -----
    def inspect
      "<#{self.class} Name:#{self.Name.inspect}>"
    end
  end
  
  # ====================================================
  # ================= NotesACL Class ===============
  # ====================================================
  class NotesACL < GripWrapper
    def Parent
      NotesDatabase.new(@raw_object.Parent)
    end
    
    def CreateACLEntry( name, level )
      raw_ACLEntry = @raw_object.CreateACLEntry( name, level )
      raw_ACLEntry ? NotesACLEntry.new(raw_ACLEntry) : nil
    end
    
    def GetEntry( name )
      raw_ACLEntry = @raw_object.GetEntry( name )
      raw_ACLEntry ? NotesACLEntry.new(raw_ACLEntry) : nil
    end
    
    def GetFirstEntry
      raw_ACLEntry = @raw_object.GetFirstEntry()
      raw_ACLEntry ? NotesACLEntry.new(raw_ACLEntry) : nil
    end
    
    def GetNextEntry( entry )
      raw_ACLEntry = @raw_object.GetNextEntry(toRaw(entry))
      raw_ACLEntry ? NotesACLEntry.new(raw_ACLEntry) : nil
    end
    
    def RemoveACLEntry( name )
      @raw_entry.RemoveACLEntry( name )
    end
    
    def RenameRole( oldName, newName )
      @raw_object.RenameRole( oldName, newName )
    end
    
    # ----- Additional Methods ----
    def each_entry
      raw_entry = @raw_object.GetFirstEntry
      while raw_entry
        next_entry = @raw_object.GetNextEntry(raw_entry)
        yield NotesACLEntry.new(raw_entry)
        raw_entry = next_entry
      end
    end
    alias each each_entry
  end
  
  # ====================================================
  # ============= NotesACLEntry Class ==================
  # ====================================================
  class	NotesACLEntry < GripWrapper
    def NameObject()
      raw_nameObject = @raw_object.NameObject
      NotesName.new(raw_nameObject)
    end
    
    def Parent
      raw_aclEntry = @raw_object.Parent
      NotesACL.new(raw_aclEntry)
    end
    
    ACLTYPE_UNSPECIFIED = 0
    ACLTYPE_PERSON = 1
    ACLTYPE_SERVER = 2
    ACLTYPE_MIXED_GROUP = 3
    ACLTYPE_PERSON_GROUP = 4
    ACLTYPE_SERVER_GROUP = 5
    def UserType
      @raw_object.UserType
    end
    # ----- Additional Methods ----
    def inspect
      "<#{self.class} Name:#{self.Name}>"
    end
  end
  
  # ====================================================
  # ============= NotesOutline Class ==================
  # ====================================================
  class NotesOutline < GripWrapper
  end
  
  # ====================================================
  # ============= NotesOutlineEntry Class ==================
  # ====================================================
  class NotesOutlineEntry < GripWrapper
  end
  
  # ====================================================
  # ============= NotesReplication Class ==================
  # ====================================================
  class NotesReplication < GripWrapper
  end
  
  # ====================================================
  # ============= NotesReplicationEntry Class ==================
  # ====================================================
  class NotesReplicationEntry < GripWrapper
  end
  
  # ====================================================
  # ======== NotesNoteCollection Class ===============
  # ====================================================
  class NotesNoteCollection < GripWrapper
  end
end
