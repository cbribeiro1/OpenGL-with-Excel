This project was carried out to improve performance when reading 3D objects using the assimp library and also to read images using stb_image within C++ and by creating a DLL that can be read by 64-bit Excel. 
With this process, it is possible to read more than 30 3D file formats and, through integration with the OpenGL library, plot on a VBA or Windows form within Excel through APIs. 
This process only works for those who have 64-bit Excel because I cannot compile the library this way for 32-bit. 
If you have Visual Studio 2022, you can compile the code; if you do not, use the compiled files. 
Copy the files from the x64\Release folder (zlib1.dll, pugixml.dll, poly2tri.dll, ModelLoader.dll, minizip.dll and assimp-vc143-mt.dll) to your C:\Windows\System32 folder
Open the Excel file (x64\Release\ModelLoader.xlsb) and run the code to verify that it is working.
