module Notesgrip
  # ====================================================
  # ============= NotesDocument Class ================
  # ====================================================
  class NotesDocument < GripWrapper
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
    
    # ---- Additional Methods ------
    def unid
      @raw_object.UniversalID
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
      "<#{self.class}, Form:#{self['Form']}>"
    end
    
  end

  # ====================================================
  # ====== NotesDocumentCollection Class ===============
  # ====================================================
  class NotesDocumentCollection < GripWrapper
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

end
