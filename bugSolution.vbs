Function GetObjectInCollection(collection, key)
  For Each obj In collection
    If obj(\"key\") = key Then
      Set GetObjectInCollection = obj
      Exit For
    End If
  Next
End Function