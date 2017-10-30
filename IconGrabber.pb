; Source: http://www.purebasic.fr/english/viewtopic.php?f=12&t=55320

Structure Icon_Dimension
   Width.i     ; Width
   Height.i    ; Height
   BitCnt.a    ; Bits per pixel: 1=2 Colors, 2=4 Colors, 4=16 Colors, 8=256 Colors, 24=TrueColor, 32=TrueColor+Alpha
EndStructure

Structure Icon_Size
   Array Icon_Size.Icon_Dimension(1)
EndStructure

Structure Icon_EnumData
   Mode.i ; what kind of data we'd like to get
   Nr.i   ; which Icon we'd like to get (1=first, 2=second, ..)
   RetVal.i; return value with the requested information
   Array Size.Icon_Dimension(1)  ; size of the icon we want or it holds the sizes of all available icons
EndStructure

Structure ICONDIRENTRY
   bWidth.a       ; Width, in pixels, of the image
   bHeight.a      ; Height, in pixels, of the image
   bColorCount.a  ; Number of colors in image (0 if >=8bpp)
   bReserved.a    ; Reserved ( must be 0)
   wPlanes.w      ; Color Planes
   wBitCount.w    ; Bits per pixel
   dwBytesInRes.l ; How many bytes in this resource?
   dwImageOffset.l; Where in the file is this image?
EndStructure

Structure ICONDIR
   idReserved.w ; Reserved (must be 0)
   idType.w     ; Resource Type (1 for icons)
   idCount.w    ; How many images?
   Array idEntries.ICONDIRENTRY(1) ;An entry for each image (idCount of 'em)
EndStructure

Enumeration
   #Icon_GetSclHdl   ; Icon might be scaled (in case a 256Pixel is requested and only 16Pixel is available, it's scaled)
   #Icon_GetCnt      ; How many different Icons are available
   #Icon_GetRealHdl  ; Get real icon - nothing is scaled! In case there's no fitting Icon you'll get nothing
   #Icon_GetSizes    ; Get List of available Icon Sizes
EndEnumeration

Procedure __Icon_EnumCallback(hLibrary, lpszType, lpszName, *Data.Icon_EnumData)
   ; 20130707..nalor..implemented support for icons with low BPP (1,2,4)
   Protected hIcon.i
   Protected hRsrc.i
   Protected hGlobal.i
   Protected Cnt.i
   Protected *GrpIconDir.GRPICONDIR
   Protected *IconImage.ICONIMAGE
   Protected Colors.a   
   
   Select *Data\Mode
      Case #Icon_GetCnt
         *Data\RetVal+1
         
      Case #Icon_GetSclHdl
         *Data\RetVal+1   
         If *Data\Nr=*Data\RetVal ; we've reached the requested Icon
            *Data\RetVal=LoadImage_(hLibrary, lpszName, #IMAGE_ICON, *Data\Size(0)\Width, *Data\Size(0)\Height, 0)
            ProcedureReturn #False ; no need to continue enumerating
         EndIf
         
      Case #Icon_GetSizes
         *Data\RetVal+1   
         If *Data\Nr=*Data\RetVal ; we've reached the requested Icon
            ;Find the group resource which lists its images
            hRsrc=FindResource_(hLibrary, lpszName, lpszType)
            If hRsrc
               ;Load And Lock To get a pointer To a GRPICONDIR
               hGlobal=LoadResource_(hLibrary, hRsrc)
               If hGlobal
                  *GrpIconDir=LockResource_(hGlobal)
                  If *GrpIconDir
                     ReDim *Data\Size(*GrpIconDir\idCount-1)
                     For Cnt=0 To *GrpIconDir\idCount-1
                        *Data\Size(Cnt)\Width=PeekA(@*GrpIconDir\idEntries[Cnt]\bWidth) ; peeka is used because 'Byte' is signed and we want an unsigned result
                        *Data\Size(Cnt)\Height=PeekA(@*GrpIconDir\idEntries[Cnt]\bHeight)
                     
                        If *Data\Size(Cnt)\Height=0
                           *Data\Size(Cnt)\Height=256
                        EndIf
                        If *Data\Size(Cnt)\Width=0
                           *Data\Size(Cnt)\Width=256
                        EndIf
                        
                        Select *GrpIconDir\idEntries[Cnt]\bColorCount
                           Case 0  ; it's an icon with at least 256 colors
                              *Data\Size(Cnt)\BitCnt=*GrpIconDir\idEntries[Cnt]\wBitCount
                           Case 2
                              *Data\Size(Cnt)\BitCnt=1
                           Case 4
                              *Data\Size(Cnt)\BitCnt=2
                           Case 16
                              *Data\Size(Cnt)\BitCnt=4
                        EndSelect                     
                        
;                        Debug "Callback Cnt >"+Str(Cnt+1)+"< Size >"+Str(*Data\Size(Cnt)\Width)+" x "+Str(*Data\Size(Cnt)\Height)+"< Col >"+Str(*Data\Size(Cnt)\BitCnt)+"<"
                     Next
                  Else
                     *Data\RetVal=-3
                  EndIf
               Else
                  *Data\RetVal=-2
               EndIf
            Else
               *Data\RetVal=-1
            EndIf         
            ProcedureReturn #False ; no need to continue enumerating
         EndIf            
                  
      Case #Icon_GetRealHdl

         Select *Data\Size(0)\BitCnt
            Case 1
               Colors=2
            Case 2
               Colors=4
            Case 4
               Colors=16
            Default
               Colors=1 ; an impossible value
         EndSelect         
         
         *Data\RetVal+1
;          Debug "Current IconNr >"+Str(*Data\RetVal)+"< Dest >"+Str(*Data\Nr)+"<"
         If *Data\Nr=*Data\RetVal ; we've reached the requested Icon         
            ; http://msdn.microsoft.com/en-us/library/ms997538.aspx
            ;Find the group resource which lists its images
            
            hRsrc=FindResource_(hLibrary, lpszName, lpszType)
            If hRsrc
               ;Load And Lock To get a pointer To a GRPICONDIR
               hGlobal=LoadResource_(hLibrary, hRsrc)
               If hGlobal
                  *GrpIconDir=LockResource_(hGlobal)
                  If *GrpIconDir
                     ; Using an ID from the group, Find, Load And Lock the RT_ICON
                     *Data\RetVal=0 ; in case the requested icon is not available, "0" will be the return value
                     For Cnt=0 To *GrpIconDir\idCount-1
                        If PeekA(@*GrpIconDir\idEntries[Cnt]\bWidth)=*Data\Size(0)\Width And
                           PeekA(@*GrpIconDir\idEntries[Cnt]\bHeight)=*Data\Size(0)\Height And
                           (*GrpIconDir\idEntries[Cnt]\wBitCount=*Data\Size(0)\BitCnt Or *Data\Size(0)\BitCnt=0 Or *GrpIconDir\idEntries[Cnt]\bColorCount=Colors)
                           
                           hRsrc=FindResource_(hLibrary, *GrpIconDir\idEntries[Cnt]\nID, #RT_ICON)
                           If hRsrc
                              hGlobal=LoadResource_(hLibrary, hRsrc)
                              If hGlobal
                                 *IconImage=LockResource_(hGlobal) ;Here, *IconImage points To an ICONIMAGE Structure
                                 If *IconImage
                                    *Data\RetVal=CreateIconFromResourceEx_(*IconImage, SizeofResource_(hLibrary, hRsrc), #True, $30000, 0, 0, 0)
                                 Else
                                    *Data\RetVal=-6
                                 EndIf
                              Else
                                 *Data\RetVal=-5
                              EndIf
                           Else
                              *Data\RetVal=-4
                           EndIf
                           ProcedureReturn #False ; we found the specific icon
                        EndIf
                     Next
                  Else
                     *Data\RetVal=-3
                  EndIf
               Else
                  *Data\RetVal=-2
               EndIf
            Else
               *Data\RetVal=-1
            EndIf
            ProcedureReturn #False ; no need to continue enumerating
         EndIf               
            
   EndSelect         
   
   ProcedureReturn #True

EndProcedure

Procedure.i __Icon_GetRealHdlICO(File.s, Width.i, Height.i, BPP.a=0)
   Protected pIconDir.ICONDIR ;We need an ICONDIR To hold the Data
   Protected *IconImage.ICONIMAGE
   Protected FileHdl.i
   Protected Cnt.i
   Protected Colors.a
   Protected RetVal.i=0
   
   Select BPP
      Case 1
         Colors=2
      Case 2
         Colors=4
      Case 4
         Colors=16
      Default
         Colors=1 ; an impossible value
   EndSelect
         
   
   FileHdl=ReadFile(#PB_Any, File)
    If FileHdl
       pIconDir\idReserved=ReadWord(FileHdl) ; Read the Reserved word
       pIconDir\idType=ReadWord(FileHdl) ; Read the Type word - make sure it is 1 For icons
       If (pIconDir\idType=#IMAGE_ICON Or pIconDir\idType=#IMAGE_CURSOR) ; it's an Icon or a Cursor
          pIconDir\idCount=ReadWord(FileHdl) ; Read the count - how many images in this file?
          ReDim pIconDir\idEntries(pIconDir\idCount -1) ; Reallocate IconDir so that idEntries has enough room For idCount elements
          If ReadData(FileHdl, @pIconDir\idEntries(0), SizeOf(ICONDIRENTRY) * pIconDir\idCount) ; Read the ICONDIRENTRY elements
             RetVal=0
             For Cnt=0 To pIconDir\idCount-1
;                 Debug "CurIcon >"+Str(pIconDir\idEntries(Cnt)\bWidth)+"<>"+Str(pIconDir\idEntries(Cnt)\bHeight)+"<>"+Str(pIconDir\idEntries(Cnt)\wBitCount)+"<"
               If PeekA(@pIconDir\idEntries(Cnt)\bWidth)=Width And
                  PeekA(@pIconDir\idEntries(Cnt)\bHeight)=Height And
                  (pIconDir\idEntries(Cnt)\wBitCount=BPP Or BPP=0 Or pIconDir\idEntries(Cnt)\bColorCount=Colors)                
                
                   *IconImage=AllocateMemory(pIconDir\idEntries(Cnt)\dwBytesInRes) ; Allocate memory To hold the image
                   If *IconImage
                       FileSeek(FileHdl, pIconDir\idEntries(Cnt)\dwImageOffset) ; Seek To the location in the file that has the image
                       If ReadData(FileHdl, *IconImage, pIconDir\idEntries(Cnt)\dwBytesInRes) ;Read the image Data
                          RetVal=CreateIconFromResourceEx_(*IconImage, pIconDir\idEntries(Cnt)\dwBytesInRes, #True, $30000, 0, 0, 0)
                       Else
                          Debug "ERROR!! Reading ICONIMAGE data (__Icon_GetRealHdlICO)"
                          RetVal=-5
                       EndIf
   
                       FreeMemory(*IconImage)
                      Break;
                   Else
                      Debug "ERROR!! Allocating Memory (__Icon_GetRealHdlICO)"
                      RetVal=-4
                   EndIf
                EndIf
             Next
          Else
             Debug "ERROR!! Reading ICONDIRENTRY data (__Icon_GetRealHdlICO)"
             RetVal=-3
          EndIf
      Else
         Debug "ERROR!! it's not an icon or a cursor (__Icon_GetRealHdlICO)"
         RetVal=-2
      EndIf
      CloseFile(FileHdl)
   Else
      Debug "ERROR!! reading file (__Icon_GetRealHdlICO)"
      RetVal=-1
   EndIf
   
   ProcedureReturn RetVal
EndProcedure

Procedure.i __Icon_GetSizesICO(File.s, *Sizes.Icon_Size)
   ; 20130707..nalor..implemented support for icons with low BPP (1,2,4)   
   Protected pIconDir.ICONDIR ;We need an ICONDIR To hold the Data
   Protected FileHdl.i
   Protected Cnt.i
   Protected RetVal.i=0
   
   FileHdl=ReadFile(#PB_Any, File)
    If FileHdl
       pIconDir\idReserved=ReadWord(FileHdl) ; Read the Reserved word
       pIconDir\idType=ReadWord(FileHdl) ; Read the Type word - make sure it is 1 For icons
       If (pIconDir\idType=#IMAGE_ICON Or pIconDir\idType=#IMAGE_CURSOR) ; it's an Icon or a Cursor
          pIconDir\idCount=ReadWord(FileHdl) ; Read the count - how many images in this file?
          ReDim pIconDir\idEntries(pIconDir\idCount -1) ; Reallocate IconDir so that idEntries has enough room For idCount elements
          If ReadData(FileHdl, @pIconDir\idEntries(0), SizeOf(ICONDIRENTRY) * pIconDir\idCount) ; Read the ICONDIRENTRY elements
             
             ReDim *Sizes\Icon_Size(pIconDir\idCount -1)
             For Cnt=0 To pIconDir\idCount-1
               *Sizes\Icon_Size(Cnt)\Width=PeekA(@pIconDir\idEntries(Cnt)\bWidth) ; peeka is used because 'Byte' is signed and we want an unsigned result
               *Sizes\Icon_Size(Cnt)\Height=PeekA(@pIconDir\idEntries(Cnt)\bHeight)
               
               If *Sizes\Icon_Size(Cnt)\Width=0
                  *Sizes\Icon_Size(Cnt)\Width=256
               EndIf
               If *Sizes\Icon_Size(Cnt)\Height=0
                  *Sizes\Icon_Size(Cnt)\Height=256
               EndIf                  
               
               Select pIconDir\idEntries(Cnt)\bColorCount
                  Case 0  ; it's an icon with at least 256 colors
                     *Sizes\Icon_Size(Cnt)\BitCnt=pIconDir\idEntries(Cnt)\wBitCount
                  Case 2
                     *Sizes\Icon_Size(Cnt)\BitCnt=1
                  Case 4
                     *Sizes\Icon_Size(Cnt)\BitCnt=2
                  Case 16
                     *Sizes\Icon_Size(Cnt)\BitCnt=4
               EndSelect               
            
               Debug "__Icon_GetSizesICO Cnt >"+Str(Cnt+1)+"/"+Str(pIconDir\idCount)+"< Size >"+Str(*Sizes\Icon_Size(Cnt)\Width)+" x "+Str(*Sizes\Icon_Size(Cnt)\Height)+"< BitCnt >"+Str(*Sizes\Icon_Size(Cnt)\BitCnt)+"<"
 
            Next
            RetVal=1
          Else
             Debug "ERROR!! Reading ICONDIRENTRY data (__Icon_GetSizesICO)"
             RetVal=-3
          EndIf
      Else
         Debug "ERROR!! it's not an icon or a cursor (__Icon_GetSizesICO)"
         RetVal=-2
      EndIf
      CloseFile(FileHdl)
   Else
      Debug "ERROR!! reading file (__Icon_GetSizesICO)"
      RetVal=-1
   EndIf
   
   ProcedureReturn RetVal
EndProcedure

Procedure.i Icon_GetSizes(File.s, *Sizes.Icon_Size, IconNr.i=1)
   Protected IconData.Icon_EnumData
   Protected hLibrary.i
   Protected ImgHdl.i
   
   IconData\RetVal=0

   Select LCase(GetExtensionPart(File))
      Case "ico", "cur"
         IconData\RetVal=__Icon_GetSizesICO(File, *Sizes)
         Debug "Sizes >"+ArraySize(*Sizes\Icon_Size())+"<"
         Debug "Width >"+Str(*Sizes\Icon_Size(0)\Height)+"<"
         
      Case "bmp"
         ImgHdl=LoadImage(#PB_Any, File)
         If ImgHdl
            ReDim *Sizes\Icon_Size(0)
            *Sizes\Icon_Size(0)\Height=ImageHeight(ImgHdl)
            *Sizes\Icon_Size(0)\Width=ImageWidth(ImgHdl)
            *Sizes\Icon_Size(0)\BitCnt=ImageDepth(ImgHdl, #PB_Image_OriginalDepth)
            FreeImage(ImgHdl)
         Else
            Debug "ERROR!! Loading File (Icon_GetSizes)"
            IconData\RetVal=-1
         EndIf
         
      Case "exe", "dll"
         hLibrary = LoadLibraryEx_(File, #Null, #LOAD_LIBRARY_AS_DATAFILE)
         If hLibrary
            IconData\Mode=#Icon_GetSizes
            IconData\Nr=IconNr
            EnumResourceNames_(hLibrary, #RT_GROUP_ICON, @__Icon_EnumCallback(), @IconData)
            FreeLibrary_(hLibrary)
            
            If IconData\RetVal ; detection of sizes succesfull
               If CopyArray(IconData\Size(), *Sizes\Icon_Size())
                  IconData\RetVal=#True
               Else
                  Debug "Error CopyArray"
                  IconData\RetVal=#False
               EndIf
            Else
               IconData\RetVal=#False
               Debug "Error Callback (Icon_GetSizes)"
            EndIf                        
         EndIf            

   EndSelect
   
   ProcedureReturn IconData\RetVal
EndProcedure

Procedure.i Icon_GetHdl(File.s, Width.i=16, IconNr.i=1, Height.i=0)
   Protected IconData.Icon_EnumData
   Protected hLibrary.i
   
   If Height=0
      Height=Width
   EndIf

   Select LCase(GetExtensionPart(File))
      Case "ico"
         IconData\RetVal=LoadImage_(#Null, @File, #IMAGE_ICON, Width, Height, #LR_LOADFROMFILE)
         
      Case "cur"   
         IconData\RetVal=LoadImage_(#Null, @File, #IMAGE_CURSOR, Width, Height, #LR_LOADFROMFILE)
      
      Case "bmp"
         IconData\RetVal=LoadImage_(#Null, @File, #IMAGE_BITMAP, Width, Height, #LR_LOADFROMFILE)
         
      Case "exe", "dll"
         hLibrary = LoadLibraryEx_(File, #Null, #LOAD_LIBRARY_AS_DATAFILE)
         If hLibrary
            IconData\RetVal=0
            IconData\Size(0)\Width=Width
            IconData\Size(0)\Height=Height
            IconData\Nr=IconNr
            IconData\Mode=#Icon_GetSclHdl
            EnumResourceNames_(hLibrary, #RT_GROUP_ICON, @__Icon_EnumCallback(), @IconData)
            FreeLibrary_(hLibrary)
         EndIf
      Default
         IconData\RetVal=0
   EndSelect
   
   ProcedureReturn IconData\RetVal
EndProcedure

Procedure.i Icon_GetRealHdl(File.s, Width.i=16, IconNr.i=1, Height.i=0, BPP.a=0)
   Protected IconData.Icon_EnumData
   Protected hLibrary.i
   Protected ImgHdl.i
   
   If Width=256 And LCase(GetExtensionPart(File))<>"bmp" ; this rule is not for BMP files!
      Width=0 ; Width256 is stored as 0 (because it's only 1 Byte, so 255 is max.)
   EndIf
   
   If Height=0 Or Height=256
      Height=Width
   EndIf
   
   IconData\RetVal=0

   Select LCase(GetExtensionPart(File))
      Case "ico", "cur"
         IconData\RetVal=__Icon_GetRealHdlICO(File, Width, Height, BPP)
         
      Case "bmp"
         ImgHdl=LoadImage(#PB_Any, File)
         If ImgHdl
            If ImageHeight(ImgHdl)=Height And
               ImageWidth(ImgHdl)=Width And
               (ImageDepth(ImgHdl, #PB_Image_OriginalDepth)=BPP Or BPP=0)
               IconData\RetVal=Icon_GetHdl(File, Width, 1, Height)
            Else
               Debug "BMP size is not available"
               IconData\RetVal=0
            EndIf
            FreeImage(ImgHdl)
         Else
            Debug "ERROR!! Loading File (Icon_GetSizes)"
            IconData\RetVal=-1
         EndIf         
         
         
      Case "exe", "dll"
         hLibrary = LoadLibraryEx_(File, #Null, #LOAD_LIBRARY_AS_DATAFILE)
         If hLibrary
            IconData\Mode=#Icon_GetRealHdl
            IconData\Nr=IconNr
            IconData\Size(0)\Width=Width
            IconData\Size(0)\Height=Height
            IconData\Size(0)\BitCnt=BPP
            EnumResourceNames_(hLibrary, #RT_GROUP_ICON, @__Icon_EnumCallback(), @IconData)
            FreeLibrary_(hLibrary)
         EndIf            
         
   EndSelect   
   
   ProcedureReturn IconData\RetVal
EndProcedure

Procedure.i Icon_GetCnt(File.s)
   Protected IconData.Icon_EnumData
   Protected hLibrary.i
   
   Select LCase(GetExtensionPart(File))
      Case "exe", "dll"   
         IconData\Mode=#Icon_GetCnt
         IconData\RetVal=0
         
         hLibrary = LoadLibraryEx_(File, #Null, #LOAD_LIBRARY_AS_DATAFILE)
         If hLibrary
            EnumResourceNames_(hLibrary, #RT_GROUP_ICON, @__Icon_EnumCallback(), @IconData)
            FreeLibrary_(hLibrary)
         EndIf
      Case "ico", "cur", "bmp"
         IconData\RetVal=1
         
      Default
         IconData\RetVal=-1 ; File not supported
   EndSelect
   
   ProcedureReturn IconData\RetVal
EndProcedure

Procedure Icon_DestroyHdl(hIcon.i)
   DestroyIcon_(hIcon)
EndProcedure

Procedure Icon_GetInfo(hIcon.i, *Size.Icon_Dimension)
   ; http://stackoverflow.com/questions/1913468/how-to-determine-the-size-of-an-icon-from-a-hicon
   Protected IconInf.ICONINFO
   Protected BMInf.BITMAP
   Protected RetVal.i=1
   
   If (GetIconInfo_(hIcon, IconInf))
      
      If (IconInf\hbmColor) ; ' Icon has colour plane
         If (GetObject_(IconInf\hbmColor, SizeOf(BITMAP), @BMInf))
            *Size\Width = BMInf\bmWidth
            *Size\Height = BMInf\bmHeight
            *Size\BitCnt = BMInf\bmBitsPixel
            DeleteObject_(IconInf\hbmColor)
         Else
            RetVal=-3
            Debug "ERROR!! GetObject failed (Icon_GetInfo)"
         EndIf
      Else ;' Icon has no colour plane, image data stored in mask
         If (GetObject_(IconInf\hbmMask, SizeOf(BITMAP), @BMInf))
            *Size\Width = BMInf\bmWidth
            *Size\Height = BMInf\bmHeight / 2
            *Size\BitCnt = 1
            DeleteObject_(IconInf\hbmMask)
         Else
            RetVal=-2
            Debug "ERROR!! GetObject failed (Icon_GetInfo)"
         EndIf
      EndIf
   Else
      RetVal=-1
      Debug "ERROR! GetIconInfo failed (Icon_GetInfo)"
   EndIf
   
   ;Debug "Width >"+Str(*Size\Width)+"< Height >"+Str(*Size\Height)+"< BPP >"+Str(*Size\BitCnt)+"<"
   
   ProcedureReturn RetVal
EndProcedure

; IDE Options = PureBasic 5.60 (Windows - x86)
; Folding = --
; EnableXP
; CompileSourceDirectory