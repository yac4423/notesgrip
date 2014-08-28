module Notesgrip
  # ====================================================
  # ================= NotesItem Class ==================
  # ====================================================
  class NotesItem < GripWrapper
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
        # ? ‚Ç‚¤‚â‚Á‚ÄNotesDateTime‚© NotesDateRange‚©‹æ•Ê‚·‚é‚ñ‚¾?
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
end
