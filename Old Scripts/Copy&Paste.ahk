#SingleInstance force
SendMode Input
handleClip("clear")
handleClip(action)
{
   static
  
   if (action = "save")
   {
      AddNextNum:= AddNextNum < 30 ? AddNextNum+1 : 1
      HighestNum+= HighestNum < 30 ? 1 : 0
      GetNextNum:= AddNextNum , ClipArray%AddNextNum% := Clipboard
   }
   else if ((action = "get") || (action = "roll")) && (GetNextNum != 0)
   {
      if (action = "roll")
         Send ^z
      Clipboard := ClipArray%GetNextNum%
      GetNextNum:= GetNextNum > 1 ? GetNextNum-1 : HighestNum  
      Send ^v
   }
   else if ((action = "rollforward") && (GetNextNum != 0))
   {
      Send, ^z
      GetNextNum:= GetNextNum < HighestNum ? GetNextNum+1 : 1
      Clipboard:= ClipArray%GetNextNum%
      Send, ^v
   }
   else if (action = "clear")
      GetNextNum:=0, AddNextNum=HighestNum=0
}
^c::
   suspend on
   Send ^c
   suspend off
   handleClip("save")
return
#0::handleClip("clear")
#v::handleClip("get")
#\::handleClip("roll")
#^\::handleClip("rollforward")