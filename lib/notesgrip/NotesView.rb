module Notesgrip
  # ====================================================
  # ================ NotesView Class ===================
  # ====================================================
  class NotesView < GripWrapper
    include DocCollection
    
    def AllEntries
      NotesViewEntryCollection.new(@raw_object.AllEntries)
    end
    
    def ColumnNames
      #@raw_object.ColumnNames
      raise "ColumnNames() is not work."
    end
    
    def Columns
      ret_list = []
      columns = @raw_object.Columns
      return [] unless columns
      columns.each {|raw_column|
        ret_list.push NotesViewColumn.new(raw_column)
      }
      ret_list
    end
    
    def Parent
      NotesDatabase.new(@raw_object.Parent)
    end
    
    def UniversalID
      @raw_object.UniversalID
    end
    
    
    def CopyColumn( sourceColumn, destinationIndex=nil )
      #raw_object = toRaw(sourceColumn)
      #raw_viewColumn = @raw_object.CopyColumn(sourceColumn, destinationIndex)
      #raw_viewColumn ? NotesViewColumn.new(raw_viewColumn) : nil
      raise "CopyColumn() is not work."
    end
    
    def CreateColumn( position=nil, columnName=nil, formula=nil )
      #raw_viewColumn = @raw_object.CreateColumn( position, columnName, formula )
      #raw_viewColumn ? NotesViewColumn.new(raw_viewColumn) : nil
      raise "CreateColumn() is not work."
    end
    
    def CreateViewNav( cacheSize=128 )
      raw_viewNavigator = @raw_object.CreateViewNav( cacheSize )
      raw_viewNavigator ? NotesViewNavigator.new(raw_viewNavigator) : nil
    end
    
    def CreateViewNavFrom( navigatorObject, cacheSize=128 )
      raw_entry = toRaw(navigatorObject)
      raw_viewNavigator = @raw_object.CreateViewNavFrom( raw_entry, cacheSize)
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
    
    def FTSearch( query, maxDocs=0 )
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
      #raw_viewColumn = @raw_object.GetColumn(columnNumber)
      #NotesViewColumn.new(raw_viewColumn)
      raise "GetColumn() is not work."
    end
    
    def GetDocumentByKey( keyArray, exactMatch=false )
      raw_doc = @raw_object.GetDocumentByKey( keyArray, exactMatch )
      raw_doc ? NotesDocument.new(raw_doc) : nil
    end
    
    def GetEntryByKey( keyArray, exactMatch=false )
      raw_viewEntry = @raw_object.GetEntryByKey( keyArray, exactMatch)
      NotesViewEntry.new(raw_viewEntry)
    end
    
    def GetFirstDocument
      raw_doc = @raw_object.GetFirstDocument()
      raw_doc ? NotesDocument.new(raw_doc) : nil
    end
    
    def GetLastDocument
      raw_doc = @raw_object.GetLastDocument()
      raw_doc ? NotesDocument.new(raw_doc) : nil
    end
    
    def GetNextDocument(document)
      raw_doc = toRaw(document)
      raw_nextDoc = @raw_object.GetNextDocument(raw_doc)
      raw_nextDoc ? NotesDocument.new(raw_nextDoc) : nil
    end
    
    def GetNextSibling( document )
      raw_doc = toRaw(document)
      raw_nextDoc = @raw_object.GetNextSibling(raw_doc)
      raw_nextDoc ? NotesDocument.new(raw_nextDoc) : nil
    end
    
    def GetNthDocument(index)
      raw_doc = @raw_object.GetNthDocument(index)
      raw_doc ? NotesDocument.new(raw_doc) : nil
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
    
    # ------ Additional Methods -----
    def each_entry
      self.AllEntries.each {|entry|
        yield entry
      }
    end
    
    def each_column
      self.Columns.each {|column|
        yield column
      }
    end
    
    alias UNID UniversalID
    
    def count
      self.AllEntries.Count
    end
    alias size count
    
    def inspect
      "<#{self.class}, Name:#{self.Name.inspect}>"
    end
  end

  # ====================================================
  # ============= NotesViewColumn Class ================
  # ====================================================
  class NotesViewColumn < GripWrapper
    # ------ Additional Methods -----
    def inspect
      "<#{self.class}, Title:#{self.Title.inspect}>"
    end
    
    def to_s
      @raw_object.Title
    end
  end
  
  # ====================================================
  # ============= NotesViewEntry Class ================
  # A view entry represents a row in a view.
  # ====================================================
  class NotesViewEntry < GripWrapper
    def Document
      raw_doc = @raw_object.Document
      NotesDocument.new(raw_doc)
    end
    
    def UniversalID
      @raw_object.UniversalID
    end
    # ------ Additional Methods -----
    alias UNID UniversalID
    def inspect
      colValues = self.ColumnValues[0,3]
      "<#{self.class}, #{colValues.inspect}>"
    end
    
  end

  # ====================================================
  # ============= NotesViewEntryCollection Class ================
  # ====================================================
  class NotesViewEntryCollection < GripWrapper
    def each_viewEntry
      raw_entry = @raw_object.GetFirstEntry
      while raw_entry
        next_entry = @raw_object.GetNextEntry(raw_entry)
        yield NotesViewEntry.new(raw_entry)
        raw_entry = next_entry
      end
    end
    alias each each_viewEntry
    
    def count
      @raw_object.Count
    end
    alias size count
    
    def inspect
      "<#{self.class}, Count:#{self.Count}>"
    end
  end
  
  # ====================================================
  # ============= NotesViewNavigator Class ================
  # ====================================================
  class NotesViewNavigator < GripWrapper
    def GetChild(entry)
      raw_entry = toRaw(entry)
      raw_viewEntry = @raw_object.GetChild(raw_entry)
      raw_viewEntry ? NotesViewEntry.new(raw_viewEntry) : nil
    end
    
    def GetCurrent()
      raw_viewEntry = @raw_object.GetCurrent()
      raw_viewEntry ? NotesViewEntry.new(raw_viewEntry) : nil
    end
    
    def GetEntry( entry )
      raw_entry = toRaw(entry)
      raw_viewEntry = @raw_object.GetEntry(raw_entry)
      raw_viewEntry ? NotesViewEntry.new(raw_viewEntry) : nil
    end
    
    def GetFirst
      raw_viewEntry = @raw_object.GetFirst()
      raw_viewEntry ? NotesViewEntry.new(raw_viewEntry) : nil
    end
    
    def GetFirstDocument
      raw_viewEntry = @raw_object.GetFirstDocument()
      raw_viewEntry ? NotesViewEntry.new(raw_viewEntry) : nil
    end
    
    def GetLast
      raw_viewEntry = @raw_object.Getlast()
      raw_viewEntry ? NotesViewEntry.new(raw_viewEntry) : nil
    end
    
    def GetLastDocument
      raw_viewEntry = @raw_object.GetLastDocument()
      raw_viewEntry ? NotesViewEntry.new(raw_viewEntry) : nil
    end
    
    def GetNext(entry)
      raw_entry = toRaw(entry)
      raw_viewEntry = @raw_object.GetNext(raw_entry)
      raw_viewEntry ? NotesViewEntry.new(raw_viewEntry) : nil
    end
    
    def GetNextCategory(entry)
      raw_entry = toRaw(entry)
      raw_viewEntry = @raw_object.GetNextCategory(raw_entry)
      raw_viewEntry ? NotesViewEntry.new(raw_viewEntry) : nil
    end
    
    def GetNextDocument(entry)
      raw_entry = toRaw(entry)
      raw_viewEntry = @raw_object.GetNextDocument(raw_entry)
      raw_viewEntry ? NotesViewEntry.new(raw_viewEntry) : nil
    end
    
    def GetNextSibling(entry)
      raw_entry = toRaw(entry)
      raw_viewEntry = @raw_object.GetNextSibling(raw_entry)
      raw_viewEntry ? NotesViewEntry.new(raw_viewEntry) : nil
    end
    
    def GetNth( index )
      raw_viewEntry = @raw_object.GetNth( index )
      raw_viewEntry ? NotesViewEntry.new(raw_viewEntry) : nil
    end
    
    def GetParent( entry )
      raw_entry = toRaw(entry)
      raw_viewEntry = @raw_object.GetParent(raw_entry)
      raw_viewEntry ? NotesViewEntry.new(raw_viewEntry) : nil
    end
    
    def GetPos( position, separator )
      raw_viewEntry = @raw_object.GetPos( position, separator )
      raw_viewEntry ? NotesViewEntry.new(raw_viewEntry) : nil
    end
    
    def GetPrev( entry )
      raw_entry = toRaw(entry)
      raw_viewEntry = @raw_object.GetPrev(raw_entry)
      raw_viewEntry ? NotesViewEntry.new(raw_viewEntry) : nil
    end
    
    def GetPrevCategory( entry )
      raw_entry = toRaw(entry)
      raw_viewEntry = @raw_object.GetPrevCategory(raw_entry)
      raw_viewEntry ? NotesViewEntry.new(raw_viewEntry) : nil
    end
    
    def GetPrevDocument( entry )
      raw_entry = toRaw(entry)
      raw_viewEntry = @raw_object.GetPrevDocument(raw_entry)
      raw_viewEntry ? NotesViewEntry.new(raw_viewEntry) : nil
    end
    
    def GetPrevSibling( entry )
      raw_entry = toRaw(entry)
      raw_viewEntry = @raw_object.GetPrevSibling(raw_entry)
      raw_viewEntry ? NotesViewEntry.new(raw_viewEntry) : nil
    end
    
    def GotoChild( entry )
      raw_entry = toRaw(entry)
      @raw_object.GotoChild(raw_entry)
    end
    
    def GotoEntry( entry )
      raw_entry = toRaw(entry)
      @raw_object.GotoEntry(raw_entry)
    end
    
    def GotoFirst( )
      @raw_object.GotoFirst()
    end
    
    def GotoFirstDocument( )
      @raw_object.GotoFirstDocument( )
    end
    
    def GotoLast( )
      @raw_object.GotoLast( )
    end
    
    def GotoNext( entry )
      raw_entry = toRaw(entry)
      @raw_object.GotoNext(raw_entry)
    end
    
    def GotoNextCategory( entry )
      raw_entry = toRaw(entry)
      @raw_object.GotoNextCategory(raw_entry)
    end
    
    def GotoNextDocument( entry )
      raw_entry = toRaw(entry)
      @raw_object.GotoNextDocument(raw_entry)
    end
    
    def GotoNextSibling( entry )
      raw_entry = toRaw(entry)
      @raw_object.GotoNextSibling(raw_entry)
    end
    
    def GotoParent( entry )
      raw_entry = toRaw(entry)
      @raw_object.GotoParent(raw_entry)
    end
    
    def GotoPos( pos, separator )
      @raw_object.GotoPos( pos, separator )
    end
    
    def GotoPrev( entry )
      raw_entry = toRaw(entry)
      @raw_object.GotoPrev(raw_entry)
    end
    
    def GotoPrevCategory( entry )
      raw_entry = toRaw(entry)
      @raw_object.GotoPrevCategory(raw_entry)
    end
    
    def GotoPrevDocument( entry )
      raw_entry = toRaw(entry)
      @raw_object.GotoPrevDocument(raw_entry)
    end
    
    def GotoPrevSibling( entry )
      raw_entry = toRaw(entry)
      @raw_object.GotoPrevSibling(raw_entry)
    end
    
    def MarkAllRead( username=nil )
      @raw_object.MarkAllRead( username )
    end
    
    def MarkAllUnread( username=nil )
      @raw_object.MarkAllUnread( username )
    end
    
    # ----- Additional Methods -----
    def inspect
      "<#{self.class}, Count:#{self.Count}>"
    end
  end
  
end
