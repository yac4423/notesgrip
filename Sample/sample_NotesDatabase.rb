require 'notesgrip'

ns = Notesgrip::NotesSession.new
db = ns.database("technotes", "names.nsf")
p db # <Notesgrip::NotesDatabase, Name:"tech's Directory", FilePath:"names.nsf">

# Scan all documents in database
db.each_document {|doc|
  p doc
}

# Scan all views in database
db.each_view {|view|
  p view
}

# Get View
view = db.view("People")
p view

# Scan all form in database
db.each_form {|form|
  p form
}

# Database Search
docList = db.FTSearch("[LastName] CONTAINS tech-notes")
p docList
docList.each {|doc|
  puts "-----"
  p doc
  doc.each_item {|item|
    p item
  }
}

# Read/Unread Document Collection
readDocList = db.GetAllReadDocuments()
unreadDocList = db.GetAllUnreadDocuments()
puts "readDocList.size = #{readDocList.size}"
puts "unreadDocList.size = #{unreadDocList.size}"

# get document by ID/UNID
doc = db.GetDocumentByID("12AE")
p doc
doc = db.GetDocumentByUNID("F8B3C89EA5A15408852572E500656D34")
p doc

# Get Form
form = db.GetForm("person")
p form

# Get Option
p db.GetOption(Notesgrip::NotesDatabase::DBOPT_SOFTDELETE)

# ACL
acl = db.ACL
acl.each_entry {|entry|
  p entry
}

# ACL Activity
db.ACLActivityLog.each {|log|
  p log
}

# Agents
db.each_agent {|agent|
  p agent
}

# Profile Collection
db.each_profile {|doc|
  p doc
}


