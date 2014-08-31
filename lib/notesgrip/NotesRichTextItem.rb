module Notesgrip
  # ====================================================
  # =========== NotesRichTextItem Class ================
  # ====================================================
  class NotesRichTextItem < NotesItem
    EMBED_ATTACHMENT = 1454
    EMBED_OBJECT = 1453
    EMBED_OBJECTLINK = 1452
    def EmbeddedObjects
      obj_list = []
      return obj_list unless @raw_object.EmbeddedObjects
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
  class NotesEmbeddedObject < GripWrapper
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
  # ================= NotesRichTextRange Class ===============
  # Represents a range of elements in a rich text item.
  # ====================================================
  class NotesRichTextRange < GripWrapper
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
  
  class NotesRichTextSection < GripWrapper
  end
  
  class NotesRichTextStyle < GripWrapper
  end
  
  class NotesRichTextTab < GripWrapper
  end
  
  # ====================================================
  # ================= NotesRichTextTable Class ===============
  # Represents a table in a rich text item.
  # ====================================================
  class NotesRichTextTable < GripWrapper
  end
  
  class NotesRichTextDocLink < GripWrapper
  end
  
  class NotesRichTextNavigator < GripWrapper
  end
  
  class NotesRichTextParagraphStyle < GripWrapper
  end
  
  class NotesMIMEEntity < GripWrapper
  end
  
  class NotesMIMEHeader < GripWrapper
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
end
