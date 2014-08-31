require 'notesgrip'

ns = Notesgrip::NotesSession.new
db = ns.database("technotes", "names.nsf")

view = db.view("People")
p view
p view.Aliases

# Column
view.each_column {|column|
  p column
}

# ViewEntry
view.each_entry {|entry|
  p entry
}

# ViewNavigator
viewNav = view.CreateViewNav()
p viewNav

# FTSearch
match_count = view.FTSearch("miwa_dankichi.tech-notes.dyndns.org")
puts "match_count = #{match_count}"
puts "After FTSearch, view.size = #{view.size}"
view.each_document {|doc|
  p doc
}
view.clear
puts "After clear, view.size = #{view.size}"

# GetAllDocumentsByKey
docList = view.GetAllDocumentsByKey(["Administrator"], false)
puts "docList.size = #{docList.size}"
docList.each {|doc|
  p doc
}

# GetAllEntriesByKey
entryList = view.GetAllEntriesByKey(["Administrator"], false)
puts "entryList.size = #{entryList.size}"
entryList.each {|entry|
  p entry
}

# GetAllReadEntries
entryList = view.GetAllReadEntries()
puts "GetAllReadEntries, entryList.size = #{entryList.size}"

# GetDocumentByKey
doc = view.GetDocumentByKey(["Administrator"])
p doc

# GetEntryByKey
entry = view.GetEntryByKey(["Administrator"])
p entry

# GetNthDocument
doc = view.GetNthDocument(2)
p doc['FullName']

