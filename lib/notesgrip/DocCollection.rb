module Notesgrip
  module DocCollection
    def each_document
      raw_doc = @raw_object.GetFirstDocument
      while raw_doc
        next_doc = @raw_object.GetNextDocument(raw_doc)
        yield NotesDocument.new(raw_doc)
        raw_doc = next_doc
      end
    end
    alias each each_document
    
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
end