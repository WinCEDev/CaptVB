Attribute VB_Name = "ImageListEx"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Bitmap load functions.

Public Declare Function ImageListEx_SHLoadDIBitmap _
               Lib "Coredll" _
               Alias "SHLoadDIBitmap" (ByVal szFileName As String) As Long

Public Declare Function ImageListEx_DeleteObject _
               Lib "Coredll" _
               Alias "DeleteObject" (ByVal hObject As Long) As Boolean

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ImageList functions.

Public Declare Function ImageListEx_ImageList_Create _
               Lib "Coredll" _
               Alias "ImageList_Create" (ByVal cx As Long, _
                                         ByVal cy As Long, _
                                         ByVal Flags As Long, _
                                         ByVal cInitial As Long, _
                                         ByVal cGrow As Long) As Long
               
Public Declare Function ImageListEx_ImageList_Add _
               Lib "Coredll" _
               Alias "ImageList_Add" (ByVal himl As Long, _
                                      ByVal hbmImage As Long, _
                                      ByVal hbmMask As Long) As Long

Public Declare Function ImageListEx_ImageList_AddMasked _
               Lib "Coredll" _
               Alias "ImageList_AddMasked" (ByVal himl As Long, _
                                            ByVal hbmImage As Long, _
                                            ByVal crMask As Long) As Long

Public Declare Function ImageListEx_ImageList_Replace _
               Lib "Coredll" _
               Alias "ImageList_Replace" (ByVal himl As Long, _
                                          ByVal i As Long, _
                                          ByVal hbmImage As Long, _
                                          ByVal hbmMask As Long) As Long

Public Declare Function ImageListEx_ImageList_GetImageCount _
               Lib "Coredll" _
               Alias "ImageList_GetImageCount" (ByVal himl As Long) As Long

Public Declare Function ImageListEx_ImageList_Remove _
               Lib "Coredll" _
               Alias "ImageList_Remove" (ByVal himl As Long, _
                                         ByVal i As Long) As Boolean

Public Declare Function ImageListEx_ImageList_RemoveAll _
               Lib "Coredll" _
               Alias "ImageList_RemoveAll" (ByVal himl As Long) As Boolean

Public Declare Function ImageListEx_ImageList_Destroy _
               Lib "Coredll" _
               Alias "ImageList_Destroy" (ByVal himl As Long) As Boolean

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ImageList flag values.
'https://learn.microsoft.com/en-us/previous-versions/ms960944(v=msdn.10)
Public Const ImageListEx_ILC_COLOR    As Long = &H0

Public Const ImageListEx_ILC_COLORDDB As Long = &HFE

Public Const ImageListEx_ILC_MASK     As Long = &H1

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Mask values.

Private Const ImageListEx_CLR_DEFAULT As Long = &HFF000000

Public Function ImageListEx_Replace(ByVal ImageListHandle As Long, _
                                    ByVal Index As Long, _
                                    ByVal ImagePath As String) As Long

    Dim lngBitmap As Long

    lngBitmap = ImageListEx_SHLoadDIBitmap(ImagePath)

    ImageListEx_Add = ImageListEx_ImageList_Replace(ImageListHandle, Index, ImagePath, 0)

    ImageListEx_DeleteObject lngBitmap

End Function

Public Function ImageListEx_Remove(ByVal ImageListHandle As Long, _
                                   ByVal Index As Long) As Boolean
    ImageListEx_Remove = ImageListEx_ImageList_Remove(ImageListHandle, Index)
End Function

Public Function ImageListEx_RemoveAll(ByVal ImageListHandle As Long) As Boolean
    ImageListEx_RemoveAll = ImageListEx_ImageList_RemoveAll(ImageListHandle)
End Function

Public Function ImageListEx_Add(ByVal ImageListHandle As Long, _
                                ByVal ImagePath As String) As Long

    Dim lngBitmap As Long

    lngBitmap = ImageListEx_SHLoadDIBitmap(ImagePath)

    ImageListEx_Add = ImageListEx_ImageList_Add(ImageListHandle, lngBitmap, 0)

    ImageListEx_DeleteObject lngBitmap

End Function

Public Function ImageListEx_AddMasked(ByVal ImageListHandle As Long, _
                                      ByVal ImagePath As String, _
                                      ByVal MaskColor As ColorConstants) As Long

    Dim lngBitmap As Long

    lngBitmap = ImageListEx_SHLoadDIBitmap(ImagePath)

    ImageListEx_AddMasked = ImageListEx_ImageList_AddMasked(ImageListHandle, lngBitmap, MaskColor)

    ImageListEx_DeleteObject lngBitmap

End Function

Public Function ImageListEx_Create(ByVal ImageWidth As Long, _
                                   ByVal ImageHeight As Long, _
                                   ByVal Flags As Long) As Long

    Const INITIAL_IMAGES   As Long = 0 'How many images the ImageList shoud initially contain.

    Const GROW_BY          As Long = 1 'By how many images to grow the ImageList when a new image is added.

    Dim lngImageListHandle As Long

    lngImageListHandle = ImageListEx_ImageList_Create(ImageWidth, ImageHeight, Flags, INITIAL_IMAGES, GROW_BY)

    ImageListEx_Create = lngImageListHandle
End Function

Public Function ImageListEx_Destroy(ByVal ImageListHandle As Long) As Long
    ImageListEx_Destroy = ImageListEx_ImageList_Destroy(ImageListHandle)
End Function

Public Function ImageListEx_GetImageCount(ByVal ImageListHandle As Long) As Long
    ImageListEx_GetImageCount = ImageListEx_ImageList_GetImageCount(ImageListHandle)
End Function



