# OutlookDataObject
OutlookDataObject Update to the 2008 David Ewen class to reflect the new Outlook and Outlook Classic


Example of Usage

    private void Control_DragDrop(object sender, DragEventArgs e)
    {

        string[] dropFileList = (string[])e.Data.GetData(DataFormats.FileDrop);

        // If "dropFileList" is not defined use FileGroupDescriptor approach
        if (dropFileList == null)
        {
            //wrap standard IDataObject in OutlookDataObject
            OutlookDataObject dataObject = new OutlookDataObject(e.Data);

            //get the names and data streams of the files dropped
            dropFileList = (string[])dataObject.GetData("FileGroupDescriptor");
            System.IO.MemoryStream[] filestreams = (System.IO.MemoryStream[])dataObject.GetData("FileContents");
            try
            {
                for (int fileIndex = 0; fileIndex < dropFileList.Length; fileIndex++)
                {
                    //use the fileindex to get the name and data stream
                    string path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid().ToString());
                    System.IO.Directory.CreateDirectory(path);
                    string filename = dropFileList[fileIndex];
                    string fname = System.IO.Path.Combine(path, filename);
                    System.IO.MemoryStream filestream = filestreams[fileIndex];
                    //save the file stream using its name to the application path
                    System.IO.FileStream outputStream = System.IO.File.Create(fname);
                    filestream.WriteTo(outputStream);
                    outputStream.Flush();
                    outputStream.Close();
                    AddFile(filename, fname);

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        else
        {
            // Now process the attachment as usual
            // "dropFileList" now contains a list of filenames
            foreach (string filename in dropFileList)
            {
                try
                {
                    AddFile(System.IO.Path.GetFileName(filename), filename);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }

    private void Control_DragOver(object sender, DragEventArgs e)
    {
        OutlookDataObject dataObject = new OutlookDataObject(e.Data);
        if (dataObject.GetDataPresent(DataFormats.FileDrop) || dataObject.GetDataPresent("FileGroupDescriptor")) e.Effect = DragDropEffects.Copy;

    }


Works for Copy and Paste in a very similar Manner.


  private void btnPaste_Click(object sender, EventArgs e)
  {

      try
      {
          string[] dropFileList = (string[])Clipboard.GetData(DataFormats.FileDrop);

          // If "dropFileList" is not defined use FileGroupDescriptor approach
          if (dropFileList == null)
          {
              //wrap standard IDataObject in OutlookDataObject
              OutlookDataObject dataObject = new OutlookDataObject(Clipboard.GetDataObject());

              //get the names and data streams of the files dropped
              dropFileList = (string[])dataObject.GetData("FileGroupDescriptor");
              System.IO.MemoryStream[] filestreams = (System.IO.MemoryStream[])dataObject.GetData("FileContents");
              try
              {
                  for (int fileIndex = 0; fileIndex < dropFileList.Length; fileIndex++)
                  {
                      //use the fileindex to get the name and data stream
                      string path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid().ToString());
                      System.IO.Directory.CreateDirectory(path);
                      string filename = dropFileList[fileIndex];
                      string fname = System.IO.Path.Combine(path, filename);
                      System.IO.MemoryStream filestream = filestreams[fileIndex];
                      //save the file stream using its name to the application path
                      System.IO.FileStream outputStream = System.IO.File.Create(fname);
                      filestream.WriteTo(outputStream);
                      outputStream.Flush();
                      outputStream.Close();
                      AddFile(filename, fname);

                  }
              }
              catch (Exception ex)
              {
                  MessageBox.Show(ex.Message);
              }

          }
          else
          {
              // Now process the attachment as usual
              // "dropFileList" now contains a list of filenames
              foreach (string filename in dropFileList)
              {
                  try
                  {
                      AddFile(System.IO.Path.GetFileName(filename), filename);
                  }
                  catch (Exception ex)
                  {
                    MessageBox.Show(ex.Message);
                  }
              }
          }
      }
      catch(Exception ex)
              {
                  MessageBox.Show(ex.Message);
              }
  }
