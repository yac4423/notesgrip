module Notesgrip
  # ====================================================
  # ================ NotesView Class ===================
  # ====================================================
  class NotesView < GripWrapper
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
      "<#{self.class}, Name:#{self.Name}>"
    end
  end

  # ====================================================
  # ============= NotesViewColumn Class ================
  # ====================================================
  class NotesViewColumn < GripWrapper
    def to_s
      @raw_object.Title
    end
  end

  # ====================================================
  # ============= NotesViewEntryCollection Class ================
  # ====================================================
  class NotesViewEntryCollection < GripWrapper
  end
  
  # ====================================================
  # ============= NotesViewNavigator Class ================
  # ====================================================
  class NotesViewNavigator < GripWrapper
  end
  
end
