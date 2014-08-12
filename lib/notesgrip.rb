#!ruby -Ks

# NotesGrip Ver.1.00    2014/7/12
# Copyright (C) 2014 Yac <yac@tech-notes.dyndns.org>
# You may redistribute it and/or modify it under the same license terms as Ruby.
require 'win32ole'
require 'singleton'
require 'notesgrip/NotesSession'

class WIN32OLE
  @const_defined = Hash.new 
  
  # const_load() ReDefined
  def WIN32OLE.my_const_load(ole_object, const_name_space)
    unless @const_defined[const_name_space] then
      #WIN32OLE.const_load(ole_object, const_name_space)
      @const_defined[const_name_space] = true
    end
  end
end



class FileSystemObject
  include Singleton
  def initialize
    @body =  WIN32OLE.new('Scripting.FileSystemObject')
  end
  
  def fullpath(filename)
    @body.getAbsolutePathName(filename)
  end
end




module Notesgrip
  
  # ====================================================
  # ================= NotesDatabase Class ===============
  # ====================================================
  class NotesDatabase < Notes_Wrapper
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
    
    def ACLActivityLog
      @raw_object.ACLActivityLog()
    end
    
    def Agents
      agent_arr = @raw_object.Agents
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
    
    CMPC_ARCHIVE_DELETE_COMPACT = 1 # a (アーカイブと削除の後に圧縮する)
    CMPC_ARCHIVE_DELETE_ONLY = 2    # A (圧縮せずにアーカイブと削除を実行する。a を置き換える。)
    CMPC_CHK_OVERLAP = 32768        # o と O (重複を確認する)
    CMPC_COPYSTYLE = 16             # c と C (コピースタイル。b と B を置き換える。)
    CMPC_DISABLE_DOCTBLBIT_OPTMZN = 128 # f (文書テーブルのビットマップの最適化を無効にする)
    CMPC_DISABLE_LARGE_UNKTBL = 4096    # k (サイズの大きい不明なテーブルを無効にする)
    CMPC_DISABLE_RESPONSE_INFO = 512    # H ([文書の階層情報を使用しない] を無効にする)
    CMPC_DISABLE_TRANSACTIONLOGGING = 262144 # t (トランザクションログを無効にする)
    CMPC_DISABLE_UNREAD_MARKS = 1048576      # U ([未読マークを使用しない] を無効にする)
    CMPC_DISCARD_VIEW_INDICES = 32           # d と D (ビュー索引を削除する)
    CMPC_ENABLE_DOCTBLBIT_OPTMZN = 64        # F (文書テーブルのビットマップの最適化を有効にする。f を置き換える。)
    CMPC_ENABLE_LARGE_UNKTBL = 2048          # K (サイズの大きい不明なテーブルを有効にする。k を置き換える。)
    CMPC_ENABLE_RESPONSE_INFO = 256          # h ([文書の階層情報を使用しない] を有効にする。H を置き換える。)
    CMPC_ENABLE_TRANSACTIONLOGGING = 131072  # T (トランザクションログを有効にする。t を置き換える。)
    CMPC_ENABLE_UNREAD_MARKS = 524288        # u ([未読マークを使用しない] を有効にする。U を置き換える。)
    CMPC_IGNORE_COPYSTYLE_ERRORS = 1024      # i (コピースタイルのエラーを無視する)
    CMPC_MAX_4GB = 16384                     # m と M (データベースの最大サイズを 4 ギガバイトに設定する)
    CMPC_NO_LOCKOUT = 8192                   # l と L (ユーザーを拒否しない)
    CMPC_RECOVER_INPLACE = 8                 # B (未使用領域があればそれを回復して、ファイルのサイズを削減する。b を置き換える。)
    CMPC_RECOVER_REDUCE_INPLACE = 4          # b (未使用領域を回復するが、ファイルのサイズは削減しない)
    CMPC_REVERT_FILEFORMAT = 65536           # r と R (古いファイル形式を変換しない)
    def CompactWithOptions( options )
      @raw_object.CompactWithOptions( options )
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
    
    
    FTINDEX_ALL_BREAKS = 4 # 文や段落の区切りで索引を付けます。
    FTINDEX_ATTACHED_BIN_FILES = 16 # 添付ファイル (バイナリ) に索引を付けます。
    FTINDEX_ATTACHED_FILES = 1 # 添付ファイル (生テキスト) に索引を付けます。
    FTINDEX_CASE_SENSITIVE = 8 # 大文字と小文字を区別した検索を有効にします。
    FTINDEX_ENCRYPTED_FIELDS = 2 # 暗号化されたフィールドに索引を付けます。
    def CreateFTIndex( options , recreate=false )
      @raw_object.CreateFTIndex( options , recreate )
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
    
    
    FIXUP_INCREMENTAL = 4 # 前回実行した Fixup 以降の文書のみ確認します。
    FIXUP_NODELETE = 16 # Fixup が壊れた文書を削除しないようにします。
    FIXUP_NOVIEWS = 64 # ビューを確認しません。
    FIXUP_QUICK = 2 # 迅速におおまかな確認を行います。
    FIXUP_REVERT = 32 # ID テーブルを以前のリリースの形式に戻します。
    FIXUP_TXLOGGED = 8 # トランザクションロギングが有効なデータベースを含みます。
    FIXUP_VERIFY = 1 # 変更を行いません。
    def Fixup(options)
      @raw_object.Fixup(options)
    end
    
    
    # Constant for FTSearch
    FT_DATE_ASC = 64 # 文書作成日の昇順にソートします。
    FT_DATE_DES = 32 # 文書作成日の降順にソートします。
    FT_SCORES = 8 # 適合スコアでソートします (デフォルト)。
    FT_DATABASE = 8192 # Lotus Domino データベースを含んで検索します。
    FT_FUZZY = 16384 # 関連語が検索されます。完全一致である必要はありません。
    FT_FILESYSTEM = 4096 # Lotus Domino データベースではないファイルを含んで検索します。
    FT_STEMS = 512 # 検索の基本として語幹が使用されます。
    FT_THESAURUS = 1024 # 類義語を使用して検索します。
    def FTDomainSearch( query, maxDocs, sortoptions=FT_SCORES, otheroptions=0, start=0, count=1, entryform="" )
      raw_doc = @raw_object.FTDomainSearch( query, maxDocs, sortoptions, otheroptions, start, count, entryform )
      raw_doc ? NotesDocument.new(raw_doc) : nil
    end
    
    def FTSearch( query, maxdocs, sortoptions=FT_SCORES, otheroptions=0 )
      raw_docCollection = @raw_object.FTSearch( query, maxdocs, sortoptions, otheroptions)
      NotesDocumentCollection.new(raw_docCollection)
    end
    
    def FTSearchRange( query, maxdocs, sortoptions=FT_SCORES, otheroptions=0, start=0 )
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
    
    def GetAllReadDocuments(username=nil)
      raw_docCollection = @raw_object.GetAllReadDocuments(username)
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
      raw_doc = @raw_object.GetDocumentByURL( url, reload, urllist, charset, webusername, webpassword, proxywebusername, proxywebpassword, returnimmediately )
      raw_doc ? NotesDocument.new(raw_doc) : nil
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
      raw_docCollection = @raw_object.GetModifiedDocuments( since, noteClass)
      NotesDocumentCollection.new(raw_docCollection)
    end
    
    
    DBOPT_LZCOMPRESSION = 65 # 添付ファイルに LZ1 圧縮を使用します。
    DBOPT_MAINTAINLASTACCESSED = 44 # LastAccessed プロパティを保持します。
    DBOPT_MOREFIELDS = 54 # データベースにフィールドを追加できます。
    DBOPT_NOHEADLINEMONITORS = 46 # ヘッドラインをモニターできません。
    DBOPT_NOOVERWRITE = 36 # 空きスペースに上書きしません。
    DBOPT_NORESPONSEINFO = 38 # 特定の返答階層をサポートしません。
    DBOPT_NOTRANSACTIONLOGGING = 45 # トランザクションログを無効にします。
    DBOPT_NOUNREAD = 37 # 未読マークを使用しません。
    DBOPT_OPTIMIZATION = 41 # 文書テーブルのビットマップの最適化を有効にします。
    DBOPT_REPLICATEUNREADMARKSTOANY = 71 # すべてのサーバーに未読マークを複製します。
    DBOPT_REPLICATEUNREADMARKSTOCLUSTER = 70 # クラスタサーバーにのみ未読マークを複製します。
    DBOPT_SOFTDELETE = 49 # 一時的削除を許可します。
    def GetOption(optionName)
      @raw_object.GetOption(optionName)
    end
    
    def SetOption( optionName, flag )
      @raw_object.SetOption(optionName, flag)
    end
    
    def GetOutline( outlinename )
      raw_outline = @raw_object.GetOutline( outlinename )
      NotesOutline.new(raw_outline)
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
    
    def Search( formula, notesDateTime, maxDocs )
      raw_docCollection = @raw_object.Search( formula, notesDateTime, maxDocs )
      NotesDocumentCollection.new(raw_docCollection)
    end
    
    def UnprocessedFTSearch(query, maxdocs, sortoptions=nil, otheroptions=nil )
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
    
    def inspect
      "#{self.class} #{self.name}, #{self.FilePath}"
    end
  end
  
  # ====================================================
  # ================= NotesACL Class ===============
  # ====================================================
  class NotesACL < Notes_Wrapper
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
    
    def each_entry
      raw_entry = @raw_object.GetFirstEntry
      while raw_entry
        next_entry = @raw_object.GetNextEntry(raw_entry)
        yield NotesACLEntry.new(raw_entry)
        raw_entry = next_entry
      end
    end
  end
  
  class	NotesACLEntry < Notes_Wrapper
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
  end
  
  class NotesAdministrationProcess < Notes_Wrapper
  end
  
  class NotesAgent < Notes_Wrapper
  end
  
  class NotesDateTime < Notes_Wrapper
  end
  
  class NotesForm < Notes_Wrapper
    def inspect
      "<#{self.class} Name:#{self.Name.inspect}>"
    end
  end
  
  class NotesInternational < Notes_Wrapper
  end
  
  class NotesLog < Notes_Wrapper
  end
  
  class NotesMIMEEntity < Notes_Wrapper
  end
  
  class NotesMIMEHeader < Notes_Wrapper
  end
  
  class NotesName < Notes_Wrapper
  end
  
  class NotesNewsLetter < Notes_Wrapper
  end
  
  class NotesNoteCollection < Notes_Wrapper
  end
  
  class NotesOutline < Notes_Wrapper
  end
  
  class NotesOutlineEntry < Notes_Wrapper
  end
  
  class NotesRegistration < Notes_Wrapper
  end
  
  class NotesReplication < Notes_Wrapper
  end
  
  class NotesReplicationEntry < Notes_Wrapper
  end
  
  class NotesRichTextDocLink < Notes_Wrapper
  end
  
  class NotesRichTextNavigator < Notes_Wrapper
  end
  
  class NotesRichTextParagraphStyle < Notes_Wrapper
  end
  
  # ====================================================
  # ================= NotesRichTextRange Class ===============
  # Represents a range of elements in a rich text item.
  # ====================================================
  class NotesRichTextRange < Notes_Wrapper
    def Navigator
      raw_richTextNavigator = @raw_object.Vavigator
      NotesRichTextNavigator.new(raw_richTextNavigator)
    end
    
    def Style
      raw_richTextStyle = @raw_object.Stle
      NotesRichTextStyle.new(raw_richTextStyle)
    end
    
    RTELEM_TYPE_DOCLINK = 5
    RTELEM_TYPE_FILEATTACHMENT = 8
    RTELEM_TYPE_OLE = 9
    RTELEM_TYPE_SECTION = 6
    RTELEM_TYPE_TABLE = 1
    RTELEM_TYPE_TABLECELL = 7
    RTELEM_TYPE_TEXTPARAGRAPH = 4
    RTELEM_TYPE_TEXTPOSITION = 10
    RTELEM_TYPE_TEXTRUN = 3
    RTELEM_TYPE_TEXTSTRING = 11
    def Type
      @raw_object.Type
    end
    
    def Clone
      raw_richTextRange = @raw_object.Clone
      NotesRichTextRange.new(raw_richTextRange)
    end
    
    RT_FIND_ACCENTINSENSITIVE = 4 # accent insensitive search (default is accent sensitive)
    RT_FIND_CASEINSENSITIVE = 1 # case insensitive search (default is case sensitive)
    RT_FIND_PITCHINSENSITIVE = 2 # pitch insensitive search (default is pitch sensitive)
    RT_REPL_ALL = 16    # replace all occurrences of the search string
    RT_REPL_PRESERVECASE = 8 # preserve case in the replacement string
    def FindAndReplace( target , replacemen, options=0 )
      @raw_object.FindAndReplace( target , replacemen, options )
    end
    
    def Remove
      @raw_object.Remove
      @raw_object = nil
    end
    
    def SetBegin( element )
      raw_element = toRaw(element)
      @raw_object.SetBegin(raw_element)
    end
    
    def SetEnd( element )
      raw_element = toRaw(element)
      @raw_object.SetEnd(raw_element)
    end
    
    def SetStyle( style )
      raw_style = toRaw(style)
      @raw_object.SetStyle(raw_style)
    end
  end
  
  class NotesRichTextSection < Notes_Wrapper
  end
  
  class NotesRichTextStyle < Notes_Wrapper
  end
  
  class NotesRichTextTab < Notes_Wrapper
  end
  
  # ====================================================
  # ================= NotesRichTextTable Class ===============
  # Represents a table in a rich text item.
  # ====================================================
  class NotesRichTextTable < Notes_Wrapper
    
  end
  
  class NotesStream < Notes_Wrapper
  end
  
  class NotesViewEntryCollection < Notes_Wrapper
  end
  
  class NotesViewNavigator < Notes_Wrapper
  end
  
  class NotesColorObject < Notes_Wrapper
  end
  
  # ====================================================
  # ================= NotesDbDirectory Class ===============
  # ====================================================
  class NotesDbDirectory < Notes_Wrapper
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


  module DocCollection
    def each
      raw_doc = @raw_object.GetFirstDocument
      while raw_doc
        next_doc = @raw_object.GetNextDocument(raw_doc)
        yield NotesDocument.new(raw_doc)
        raw_doc = next_doc
      end
    end
    
    def [](index)
      if index >= 0
        raw_doc = @raw_object.GetNthDocument( index + 1)  # GetNthDocument(1) is the first doc
      else
        docs_number = self.Count()
        raw_index = docs_number - index.abs
        return nil if raw_index < 0
        raw_doc = @raw_object.GetNthDocument( raw_index + 1)
      end
      raw_doc ? NotesDocument.new(raw_doc) : nil
    end
    
    def GetFirstDocument
      self[0]
    end
    def GetLastDocument
      raw_doc = @raw_object.GetLastDocument
      raw_doc ? NotesDocument.new(raw_doc) : nil
    end
    
    def GetNthDocument(index)
      if index >= 1
        self[index-1]
      else
        nil
      end
    end
    
    def GetNextDocument(document)
      raw_doc = toRaw(document)
      raw_nextDoc = @raw_object.GetNextDocument(raw_doc)
      raw_nextDoc ? NotesDocument.new(raw_nextDoc) : nil
    end
    
    def GetPrevDocument( document )
      raw_doc = toRaw(document)
      raw_prevDoc = @raw_object.GetPrevDocument(raw_doc)
      raw_prevDoc ? NotesDocument.new(raw_prevDoc) : nil
    end
    
  end

  # ====================================================
  # ====== NotesDocumentCollection Class ===============
  # ====================================================
  class NotesDocumentCollection < Notes_Wrapper
    include DocCollection
    
    def Count
      @raw_object.count
    end
    alias size Count
    
    def Parent
      NotesDatabase.new(@raw_object.Parent)
    end
    
    def AddDocument( document)
      raw_doc = toRaw(document)
      @raw_object.AddDocument(raw_doc)
    end
    
    def Clone()
       NotesDocumentCollection.new(@raw_object.Clone())
    end
    
    def Contains( inputNotes )
      raw_obj = toRaw(inputNotes)
      @raw_object.Contains(raw_obj)
    end
    
    def DeleteDocument( document )
      raw_doc = toRaw(document)
      @raw_object.DeleteDocument(raw_doc)
    end
    
    def FTSearch( query, maxDocs )
      raw_docCollection = @raw_object.FTSearch( query, maxDocs )
      FTSearch( raw_docCollection )
    end
    
    def GetDocument( document )
      raw_doc = toRaw(document)
      raw_doc2 = @raw_object.GetDocument(raw_doc)
      raw_doc2 ? NotesDocument.new(raw_doc2) : nil
    end
    
    def Intersect( inputNotes )
      raw_obj = toRaw(inputNotes)
      @raw_object.Intersect(raw_obj)
    end
    
    def Merge( inputNotes )
      raw_obj = toRaw(inputNotes)
      @raw_object.Merge(raw_obj)
    end
    
    def StampAllMulti( document )
      raw_doc = toRaw(document)
      @raw_object.StampAllMulti(raw_doc)
    end
    
    def Subtract( inputNotes )
      raw_obj = toRaw(inputNotes)
      @raw_object.Subtract(raw_obj)
    end
    
  end

  # ====================================================
  # ================ NotesView Class ===================
  # ====================================================
  class NotesView < Notes_Wrapper
    include DocCollection
    
    def AllEntries
      NotesViewEntryCollection.new(@raw_object.AllEntries)
    end
    
    def Columns
      ret_list = []
      @raw_object.Columns.each {|raw_column|
        ret_list.push NotesViewColumn.new(raw_column)
      }
      ret_list
    end
    
    def Parent
      NotesDatabase.new(@raw_object.Parent)
    end
    
    def CopyColumn( sourceColumn, destinationIndex=nil )
      raw_object = toRaw(sourceColumn)
      raw_viewColumn = @raw_object.CopyColumn(sourceColumn, destinationIndex)
      raw_viewColumn ? NotesViewColumn.new(raw_viewColumn) : nil
    end
    
    def CreateColumn( position=nil, columnName=nil, formula=nil )
      raw_viewColumn = @raw_object.CreateColumn( position, columnName, formula )
      raw_viewColumn ? NotesViewColumn.new(raw_viewColumn) : nil
    end
    
    def CreateViewNav( cacheSize )
      raw_viewNavigator = @raw_object.CreateViewNav( cacheSize )
      raw_viewNavigator ? NotesViewNavigator.new(raw_viewNavigator) : nil
    end
    
    def CreateViewNavFrom( navigatorObject, cacheSize=128 )
      raw_navi = toRaw(navigatorObject)
      raw_viewNavigator = @raw_object.CreateViewNavFrom( raw_navi, cacheSize)
      raw_viewNavigator ? NotesViewNavigator.new(raw_viewNavigator) : nil
    end
    
    def CreateViewNavFromAllUnread( username=nil )
      raw_viewNavigator = @raw_object.CreateViewNavFromAllUnread( username )
      raw_viewNavigator ? NotesViewNavigator.new(raw_viewNavigator) : nil
    end
    
    def CreateViewNavFromCategory( category, cacheSize=128 )
      raw_viewNavigator = @raw_object.CreateViewNavFromCategory( category, cacheSize)
      raw_viewNavigator ? NotesViewNavigator.new(raw_viewNavigator) : nil
    end
    
    def CreateViewNavFromDescendants( navigatorObject, cacheSize=128 )
      raw_obj = toRaw(navigatorObject)
      raw_viewNavigator = @raw_object.CreateViewNavFromDescendants( raw_obj, cacheSize )
      raw_viewNavigator ? NotesViewNavigator.new(raw_viewNavigator) : nil
    end
    
    def CreateViewNavMaxLevel( level, cacheSize=128 )
      raw_viewNavigator = @raw_object.CreateViewNavMaxLevel( level, cacheSize )
      raw_viewNavigator ? NotesViewNavigator.new(raw_viewNavigator) : nil
    end
    
    def FTSearch( query, maxDocs )
      @raw_object.FTSearch( query, maxDocs )
    end
    
    def GetAllDocumentsByKey( keyArray , exactMatch = false )
      raw_docCollection = @raw_object.GetAllDocumentsByKey( keyArray , exactMatch)
      NotesDocumentCollection.new(raw_docCollection)
    end
    
    def GetAllEntriesByKey( keyArray , exactMatch=false )
      raw_viewEntryCollection = @raw_object.GetAllEntriesByKey( keyArray , exactMatch)
      NotesViewEntryCollection.new(raw_viewEntryCollection)
    end
    
    def GetAllReadEntries( username=nil )
      raw_viewEntryCollection = @raw_object.GetAllReadEntries( username)
      NotesViewEntryCollection.new(raw_viewEntryCollection)
    end
    
    def GetAllUnreadEntries( username=nil )
      raw_viewEntryCollection = @raw_object.GetAllUnreadEntries( username)
      NotesViewEntryCollection.new(raw_viewEntryCollection)
    end
    
    def GetChild( document )
      raw_doc = toRaw(document)
      raw_respdoc = @raw_object.GetChild(raw_doc)
      raw_respdoc ? NotesDocument.new(raw_rspdoc) : nil
    end
    
    def GetColumn( columnNumber )
      raw_viewColumn = @raw_object.GetColumn(columnNumber)
      NotesViewColumn.new(raw_viewColumn)
    end
    
    def GetDocumentByKey( keyArray, exactMatch=false )
      raw_doc = @raw_object.GetDocumentByKey( keyArray, exactMatch )
      raw_doc ? NotesDocument.new(raw_doc) : nil
    end
    
    def GetEntryByKey( keyArray, exactMatch=false )
      raw_viewEntry = @raw_object.GetEntryByKey( keyArray, exactMatch)
      NotesViewEntry.new(raw_viewEntry)
    end
    
    def GetNextSibling( document )
      raw_doc = toRaw(document)
      raw_nextDoc = @raw_object.GetNextSibling(raw_doc)
      raw_nextDoc ? NotesDocument.new(raw_nextDoc) : nil
    end
    
    def GetParentDocument( document )
      raw_doc = toRaw(document)
      raw_parentDoc = @raw_object.GetParentDocument(raw_doc)
      raw_parentDoc ? NotesDocument.new(raw_parentDoc) : nil
    end
    
    def GetPrevDocument( document )
      raw_doc = toRaw(document)
      raw_prevDoc = @raw_object.GetPrevDocument(raw_doc)
      raw_prevDoc ? NotesDocument.new(raw_prevDoc) : nil
    end
    
    def GetPrevSibling( document )
      raw_doc = toRaw(document)
      raw_prevDoc = @raw_object.GetPrevSibling(raw_doc)
      raw_prevDoc ? NotesDocument.new(raw_prevDoc) : nil
    end
    
    def inspect
      "#{self.class} #{self.name}"
    end
  end

  # ====================================================
  # ============= NotesDocument Class ================
  # ====================================================
  class NotesDocument < Notes_Wrapper
    def initialize(raw_doc)
      super(raw_doc)
      @parent_db = nil
      @parent_view = nil
    end
    
    def EmbeddedObjects
      raw_embeddedObjects = @raw_object.EmbeddedObjects
      ret_list = []
      raw_embeddedObjects.each {|raw_embeddedObject|
        ret_list.push NotesEmbeddedObject.new(raw_embeddedObject)
      }
      ret_list
    end
    
    def Items
      ret_list = []
      @raw_object.Items.each {|raw_item|
        if raw_item.Type == NotesItem::RICHTEXT
          ret_list.push NotesRichTextItem.new(raw_item)
        else
          ret_list.push NotesItem.new(raw_item)
        end
      }
      ret_list
    end
    
    def ParentDatabase
      NotesDatabase.new(@raw_object.ParentDatabase)
    end
    
    def ParentView
      raw_view = @raw_object.ParentView
      raw_view ? NotesView(raw_view) : nil
    end
    
    def Responses
      raw_docCollection = @raw_object.Responses
      NotesDocumentCollection.new(raw_docCollection)
    end
    
    def UNID
      @raw_object.UniversalID
    end
    
    def AppendItemValue( itemName, value )
      raw_item = @raw_object.AppendItemValue( itemName, value )
      NotesItem.new(raw_item)
    end
    
    def AttachVCard( clientADTObject )
      raw_obj = toRaw(clientADTObject)
      @raw_object.AttachVCard(raw_obj)
    end
    
    def CopyAllItems( sourceDoc, replace=false )
      raw_sourceDoc = toRaw(sourceDoc)
      @raw_object.CopyAllItems( raw_sourceDoc, replace)
    end
    
    def CopyItem( sourceItem, newName=nil )
      raw_sourceItem = toRaw(sourceItem)
      unless newName
        newName = sourceItem.Name
      end
      raw_item = @raw_object.CopyItem(raw_sourceItem, newName)
      NotesItem.new(raw_item)
    end
    
    def CopyToDatabase( destDatabase )
      raw_destDB = toRaw(destDatabase)
      new_rawDoc = @raw_object.CopyToDatabase( raw_destDB )
      NotesDocument.new(new_rawDoc)
    end
    
    def CreateMIMEEntity( itemName )
      raw_mimeEntry = @raw_object.CreateMIMEEntity( itemName )
      NotesMIMEEntity.new(raw_mimeEntry)
    end
    
    def CreateReplyMessage( all )
      raw_replyDoc = @raw_object.CreateReplyMessage( all )
      NotesDocument.new(raw_replyDoc)
    end
    
    def CreateRichTextItem( name )
      raw_richTextItem = @raw_object.CreateRichTextItem( name )
      NotesRichTextItem.new(raw_richTextItem)
    end
    
    def GetAttachment( fileName )
      raw_embeddedObject = @raw_object.GetAttachment( fileName )
      raw_embeddedObject ? NotesEmbeddedObject.new(raw_embeddedObject) : nil
    end
    
    def GetFirstItem( name )
      raw_item = @raw_object.GetFirstItem( name )
      raw_item ? NotesItem.new(raw_item) : nil
    end
    
    def GetItemValue( itemName )
      @raw_object.GetItemValue( itemName )
    end
    
    def GetItemValueCustomDataBytes( itemName, dataTypeName )
      @raw_object.GetItemValueCustomDataBytes( itemName, dataTypeName )
    end
    
    def GetItemValueDateTimeArray( itemName )
      obj_list = @raw_object.GetItemValueDateTimeArray( itemName )
      ret_list = []
      obj_list.each {|date_obj|
        ret_list.push NotesDateTime.new(date_obj)
      }
    end
    
    def GetMIMEEntity( itemName )
      raw_mimeEntry = @raw_object.GetMIMEEntity( itemName )
      raw_mimeEntry ? NotesMIMEEntity.new(raw_mimeEntry) : nil
    end
    
    def MakeResponse( document )
      NotesDocument.new(toRaw(document))
    end
    
    def RenderToRTItem( notesRichTextItem )
      raw_richTextItem = toRaw(notesRichTextItem)
      @raw_object.RenderToRTItem(raw_richTextItem)
    end
    
    def ReplaceItemValue( itemName, value )
      raw_item = @raw_object.ReplaceItemValue( itemName, value )
      NotesItem.new(raw_item)      
    end
    
    def ReplaceItemValueCustomDataBytes( itemName, dataTypeName, byteArray )
      @raw_object.ReplaceItemValueCustomDataBytes( itemName, dataTypeName, byteArray )
    end
    
    def [](itemname)
      if @raw_object.HasItem(itemname)
        raw_item = @raw_object.GetFirstItem(itemname)
      else
        # If this document hasn't itemname field, create new item.
        raw_item = @raw_object.AppendItemValue(itemname, "")
      end
      if raw_item.Type == NotesItem::RICHTEXT
        NotesRichTextItem.new(raw_item)
      else
        NotesItem.new(raw_item)
      end
    end
    
    # Copy Field
    def []=(itemname, other_item)
      if @raw_object.HasItem(itemname)
        @raw_object.RemoveItem(itemname) # 既存のデータが入った状態でCopyItemを呼ぶと、追加処理になるので
      end
      
      raw_otherItem = toRaw(other_item)
      raw_item = @raw_object.CopyItem(raw_otherItem, itemname)
      NotesItem.new(raw_item)
    end
    
    def each_item
      self.Items.each {|notes_item|
        yield notes_item
      }
    end
    
    def save(force=true, createResponse=false , markRead = false)
      @raw_object.Save(force, createResponse, markRead)
    end
    
    def inspect
      "#{self.class} Form:#{self['Form']}"
    end
    
  end

  # ====================================================
  # ================= NotesItem Class ==================
  # ====================================================
  class NotesItem < Notes_Wrapper
    class NameError < StandardError
    end
    class RichFieldError < StandardError
    end
    
    RICH_TYP = 1
    DATE_TYP = 1024
    ATTACHMENT_TYP = 1084
    TEXT_TYP = 1280
    NAME_TYP = 1074
    
    def DateTimeValue
      raw_dateTimeVal = @raw_object.DateTimeValue
      raw_dateTimeVal ? NotesDateTime.new(raw_dateTimeVal) : nil
    end
    
    def Parent
      NotesDocument.new(@raw_object.Parent)
    end
    
    ACTIONCD = 16       # saved action CD records; non-Computable; canonical form.
    ASSISTANTINFO = 17  # saved assistant information; non-Computable; canonical form.
    ATTACHMENT = 1084   # file attachment.
    AUTHORS = 1076      # authors.
    COLLATION = 2       # COLLATION?
    DATETIMES = 1024    # date-time value or range of date-time values.
    EMBEDDEDOBJECT = 1090 # embedded object.
    ERRORITEM = 256     # an error occurred while accessing the type.
    FORMULA = 1536      # Notes formula.
    HTML = 21           # HTML source text.
    ICON = 6            # icon.
    LSOBJECT = 20       # saved LotusScript Object code for an agent.
    MIME_PART = 25      # MIME support.
    NAMES = 1074        # names.
    NOTELINKS = 7       # link to a database, view, or document.
    NOTEREFS = 4        # reference to the parent document.
    NUMBERS = 768       # number or number list.
    OTHEROBJECT = 1085  # other object.
    QUERYCD = 15        # saved query CD records; non-Computable; canonical form.
    READERS = 1075      # readers.
    RFC822Text = 1282   # RFC822 Internet mail text.
    RICHTEXT = 1        # rich text.
    SIGNATURE = 8       # signature.
    TEXT = 1280         # text or text list.
    UNAVAILABLE = 512   # the item type isn't available.
    UNKNOWN = 0         # the item type isn't known.
    USERDATA = 14       # user data.
    USERID = 1792       # user ID name.
    VIEWMAPDATA = 18    # saved ViewMap dataset; non-Computable; canonical form.
    VIEWMAPLAYOUT = 19  # saved ViewMap layout; non-Computable; canonical form.
    def Type
      @raw_object.Type
    end
    
    def Text
      @raw_object.text
    end
    
    def ValueLength
      @raw_object.ValueLength
    end
    alias length ValueLength
    
    
    def Values
      @raw_object.Values
    end
    alias values Values
    
    def Values=(item_value)
      # @raw_object.Type returns data type of field(ex: RICH_TYP,DATE_TYP...)
      case @raw_object.Type
      when RICH_TYP
        val_array = [item_value].flatten
        raw_doc = self.Parent.raw
        item_name = @raw_object.Name
        raw_doc.RemoveItem(item_name)
        new_item = raw_doc.CreateRichTextItem( item_name )
        @raw_object = new_item.AppendText(val_array.join(""))
      else
        val_array = [item_value].flatten
        raw_doc = self.Parent.raw
        item_name = @raw_object.Name
        raw_newitem = raw_doc.ReplaceItemValue(item_name, val_array)
        @raw_object = raw_newitem   # replace @raw_object
      end
    end
    alias text= Values=
    
    def AppendToTextList( newValue )
      @raw_object.AppendToTextList( newValue )
    end
    
    def contains?(value)
      @raw_object.Contains(value)
    end
    
    def CopyItemToDocument( otherDoc, newName=nil )
      unless newName
        newName = self.Name
      end
      raw_newItem = @raw_object.CopyItemToDocument( toRaw(otherDoc), newName )
      NotesItem.new(raw_newItem)
    end
    
    def GetValueDateTimeArray()
      rawDateTimeArray = @raw_object.GetValueDateTimeArray()
      return nil unless rawDateTimeArray
      ret_list = []
      rawDateTimeArray.each {|rawDateTime|
        # ? どうやってNotesDateTimeか NotesDateRangeか区別するんだ?
        ret_list.push NotesDateTime.new(rawDateTime)
        raise "Not Implement"
      }
    end
    
    def GetMIMEEntity()
      rawMIMEEntry = @raw_object.GetMIMEEntity()
      rawMIMEEntry ? NotesMIMEEntity.new(rawMIMEEntry) : nil
    end
    
    def Remove
      @raw_object.Remove()
      @raw_object = nil
    end
    
    # ----- Additional Methods ----
    
    ItemTypeDic = {
      16=>'ACTIONCD', 17=>'ASSISTANTINFO', 1084=>'ATTACHMENT', 1076=>'AUTHORS', 2=>'COLLATION', 1024=>'DATETIMES', 
      1090=>'EMBEDDEDOBJECT', 256=>'ERRORITEM', 1536=>'FORMULA', 21=>'HTML', 6=>'ICON', 20=>'LSOBJECT', 25=>'MIME_PART', 
      1074=>'NAMES', 7=>'NOTELINKS', 4=>'NOTEREFS', 768=>'NUMBERS', 1085=>'OTHEROBJECT', 15=>'QUERYCD', 1075=>'READERS', 
      1282=>'RFC822Text', 1=>'RICHTEXT', 8=>'SIGNATURE', 1280=>'TEXT', 512=>'UNAVAILABLE', 0=>'UNKNOWN', 14=>'USERDATA', 
      1792=>'USERID', 18=>'VIEWMAPDATA', 19=>'VIEWMAPLAYOUT'
    }
    def inspect
      if self.Values == nil
        val = self.text
      elsif self.Values.size > 1
        val = self.Values
      else
        val = self.Text
      end
      "<#{self.class}, Type:#{ItemTypeDic[self.Type]}, #{self.name.inspect}=>#{val.inspect}>"
    end
    
    def each_value
      @raw_object.Values.each {|val|
        yield val
      }
    end
    
    
    def richitem_attached?
      return false if @raw_object.Type == TEXT_TYP
      return true if @raw_object.EmbeddedObjects != nil
      return false
    end
    
    def to_s
      self.Text
    end
    
    def CopyItem(other_item)
      parent_doc = self.Parent
      @raw_object = parent_doc.CopyItem(other_item, self.Name)
    end
    
  end

  # ====================================================
  # ================= NotesRichTextItem Class ==================
  # ====================================================
  class NotesRichTextItem < NotesItem
    EMBED_ATTACHMENT = 1454
    EMBED_OBJECT = 1453
    EMBED_OBJECTLINK = 1452
    def EmbeddedObjects
      obj_list = []
      @raw_object.EmbeddedObjects.each {|rawEmbObj|
        obj_list.push NotesEmbeddedObject.new(rawEmbObj)
      }
      obj_list
    end
    
    def AddNewLine( n=1, forceParagraph=true )
      @raw_object.AddNewLine( n, forceParagraph)
    end
    
    def AppendDocLink( linkTo, comment=nil, hotSpotText=nil )
      raw_linkTo = toRaw(linkTo)
      @raw_object.AppendDocLink( raw_linkTo, comment, hotSpotText )
    end
    
    def AppendParagraphStyle( paragStyle )
      raw_paragStyle = toRaw(paragStyle)
      @raw_object.AppendParagraphStyle( raw_paragStyle )
    end
    
    def AppendRTItem( otherRichItem )
      raw_otherRichItem = toRaw(otherRichItem)
      @raw_object.AppendRTItem( raw_otherRichItem )
    end
    
    def AppendStyle( richTextStyle )
      raw_richTextStyle = toRaw(richTextStyle)
      @raw_object.AppendStyle( raw_richTextStyle )
    end
    
    def AppendTable( rows, columns, labels=nil, leftMargin=1440, rtpsStyleArray=nil )
      raw_styleArray = []
      if rtpsStyleArray
        rtpsStyleArray.each {|rtpsStyle|
          raw_styleArray.push toRaw(rtpsStyle)
        }
      end
      @raw_object.AppendTable( rows, columns, labels, leftMargin, raw_styleArray )
    end
    
    def AppendText( text )
      @raw_object.AppendText( text )
    end
    
    def BeginInsert( element, after=false )
      raw_element = toRaw(element)
      @raw_object.BeginInsert(raw_element)
    end
    
    def BeginSection( title, titleStyle=nil, barColor=nil, expand=false )
      raw_titleStyle = toRaw(titleStyle)
      raw_barColor = toRaw(barColor)
      @raw_object.BeginSection( title, raw_titleStyle, raw_barColor, expand)
    end
    
    def CreateNavigator()
      raw_RichTextNavigator = @raw_object.CreateNavigator()
      NotesRichTextNavigator.new(raw_RichTextNavigator)
    end
    
    def CreateRange()
      raw_notesRichTextRange = @raw_object.CreateRange()
      NotesRichTextRange.new(raw_notesRichTextRange)
    end
    
    def EmbedObject( type, appClass, source )
      case type
      when EMBED_ATTACHMENT
        fullpath = FileSystemObject.instance.fullpath(source)
        added_object = @raw_object.EmbedObject( EMBED_ATTACHMENT, "", fullpath)
      when EMBED_OBJECT
        added_object = @raw_object.EmbedObject( EMBED_OBJECT, appClass, "")
      when EMBED_OBJECTLINK
        added_object = @raw_object.EmbedObject( EMBED_OBJECTLINK, "", source)
      else
        raise "Illegal Type."
      end
      NotesEmbeddedObject.new(added_object)
    end
    
    def GetEmbeddedObject( name )
      raw_EmbeddedObject = @raw_object.GetEmbeddedObject( name )
      raw_EmbeddedObject ? NotesEmbeddedObject.new(raw_EmbeddedObject) : nil
    end
    
    def GetFormattedText( tabstrip, lineLength=0 )
      @raw_object.GetFormattedText( tabstrip, lineLength )
    end
    
    def [](embobj_name)
      raw_embobj = @raw_object.GetEmbeddedObject(embobj_name)
      if raw_embobj
        NotesEmbeddedObject.new(raw_embobj)
      else
        nil
      end
    end
    
    def add_file(filename)
      self.EmbedObject(EMBED_ATTACHMENT, "", filename)
    end
    
    def inspect
      val = self.Text
      "<#{self.class}, Type:#{ItemTypeDic[self.Type]}, #{self.name.inspect}=>#{val.inspect}>"
    end
  end

  # ====================================================
  # ============= NotesEmbeddedObject Class ================
  # An embedded object(attached file(EMBED_ATTACHMENT), object link(EMBED_OBJECTLINK), OLE(EMBED_OBJECT), etc)
  # ====================================================
  class NotesEmbeddedObject < Notes_Wrapper
    EMBED_ATTACHMENT = 1454
    EMBED_OBJECT = 1453
    EMBED_OBJECTLINK = 1452
    TypeDic = {
      1454=>'EMBED_ATTACHMENT', 1453=>'EMBED_OBJECT', 1452=>'EMBED_OBJECTLINK'
    }
    def Parent
      NotesRichTextItem.new(@raw_object.Parent)
    end
    alias parent Parent
    
    def ExtractFile( path )
      fullpath = FileSystemObject.instance.fullpath(path)
      @raw_object.ExtractFile(fullpath)
    end
    
    def Remove()
      @raw_object.Remove()
      @raw_object = nil
    end
    
    def inspect
      "<#{self.class} Type:#{TypeDic[self.Type]} Name:#{self.Name.inspect}>"
    end
    
    
  end
  
  # ====================================================
  # ============= NotesViewColumn Class ================
  # ====================================================
  class NotesViewColumn < Notes_Wrapper
    def to_s
      @raw_object.Title
    end
  end
  
end

if $0 == __FILE__
  def assert_equal(expected, val)
    puts "***** assert_equal *****"
    p expected
    p val
  end
  
  @ns = NotesSession.new
  @db = @ns.database("", "test_db.nsf")
    
  50.times {|index|
    new_doc = @db.CreateDocument("MainForm")
    new_doc["TextField"].text = sprintf("%02d", index)
    new_doc["value_field"].text = index
    new_doc["time_field"].text = Time.now
    new_doc.save
  }
  view = @db.view("keys")
  doc = view[10]
  #doc = view.raw.GetDocumentByKey(["15", 15], false)
  doc = view.get_doc_by_key("15", 15)
  p doc

end
