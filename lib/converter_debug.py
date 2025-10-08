import os
import sys
import argparse
import uno, unohelper
import time
import shlex
import glob
import base64
import json
from com.sun.star.beans import PropertyValue
from com.sun.star.connection import NoConnectException
from com.sun.star.document.UpdateDocMode import QUIET_UPDATE
from com.sun.star.lang import DisposedException, IllegalArgumentException
from com.sun.star.io import IOException, XOutputStream
from com.sun.star.script import CannotConvertException
from com.sun.star.uno import Exception as UnoException
from com.sun.star.uno import RuntimeException


desktop = None
unocontext = None
nbConsecutiveAttemptOpeningDocument = 0
nbConsecutiveAttemptOpeningDocumentMax = 10
parser = argparse.ArgumentParser()
parser.add_argument("-p", "--pipe")
parser.add_argument("-i", "--input")
parser.add_argument("-o", "--output")
parser.add_argument("-f", "--format")
parser.add_argument("-fo", "--formatOptions")


def UnoProps(**args):
    props = []
    for key in args:
        prop = PropertyValue()
        prop.Name = key
        prop.Value = args[key]
        props.append(prop)
    return tuple(props)


def send(message):
    sys.stdout.write(message)
    sys.stdout.flush()


def sendErrorOrExit(code):
    ### Tell to python that we want to modify the global variable
    global nbConsecutiveAttemptOpeningDocument
    nbConsecutiveAttemptOpeningDocument += 1
    if nbConsecutiveAttemptOpeningDocument < nbConsecutiveAttemptOpeningDocumentMax:
        send(code)  # The document could not be opened.
    else:
        send("999")  # Too many attempts, giving up.
    sys.exit()


def main():
    global desktop
    global unocontext
    global nbConsecutiveAttemptOpeningDocument
    
    if len(sys.argv) < 2:
        listen()
        return

    fileOption = parser.parse_args()
    cwd = unohelper.systemPathToFileUrl(os.getcwd())

    try:
        unocontext = uno.getComponentContext()
        resolver = unocontext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", unocontext)
        smgr = resolver.resolve("uno:pipe,name="+ fileOption.pipe +";urp;StarOffice.ServiceManager")
        remoteContext = smgr.getPropertyValue("DefaultContext")
        desktop = remoteContext.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", remoteContext)
    except NoConnectException:
        sendErrorOrExit("100")
    except UnoException:
        sendErrorOrExit("100")
    except:
        sendErrorOrExit("1")

    try:
        inputurl = unohelper.absolutize(cwd, unohelper.systemPathToFileUrl(fileOption.input))
        document = desktop.loadComponentFromURL(inputurl, "_blank", 0, UnoProps(Hidden=True, ReadOnly=True, UpdateDocMode=QUIET_UPDATE))

        if not document:
            sendErrorOrExit("400")

        nbConsecutiveAttemptOpeningDocument = 0

    except IllegalArgumentException:
        sendErrorOrExit("400")
    except DisposedException:
        sendErrorOrExit("400")
    except IOException:
        sendErrorOrExit("400")
    except CannotConvertException:
        sendErrorOrExit("400")
    except UnoException:
        sendErrorOrExit("400")
    except:
        sendErrorOrExit("1")

    # refresh indexes (required for document which include charts, indexes, etc.)
    try:
        document.refresh()
        indexes = document.getDocumentIndexes()
    except AttributeError:
        # the document doesn't implement the XRefreshable and/or
        # XDocumentIndexesSupplier interfaces
        pass
    else:
        for i in range(0, indexes.getCount()):
            indexes.getByIndex(i).update()

    outputprops = UnoProps(FilterName=fileOption.format, Overwrite=True)
    if fileOption.formatOptions != '':
        outputprops += UnoProps(FilterOptions=fileOption.formatOptions)
    outputurl = unohelper.absolutize(cwd, unohelper.systemPathToFileUrl(fileOption.output) )

    # Get output directory and filename for capturing extracted images
    output_dir = os.path.dirname(fileOption.output)
    output_basename = os.path.splitext(os.path.basename(fileOption.output))[0]
    
    # Record files in output directory before conversion
    files_before = set()
    if os.path.exists(output_dir):
        files_before = set(os.listdir(output_dir))
        send(f"DEBUG: Files before conversion: {len(files_before)} files\n")

    try:
        send(f"DEBUG: Converting to: {fileOption.format}\n")
        send(f"DEBUG: Output path: {fileOption.output}\n")
        send(f"DEBUG: Output dir: {output_dir}\n")
        send(f"DEBUG: Output basename: {output_basename}\n")
        
        document.storeToURL(outputurl, tuple(outputprops) )
        
        send(f"DEBUG: Conversion completed\n")
    except:
        sendErrorOrExit('401') # could not convert document
        document.dispose()
        document.close(True)
        return

    # Capture any new image files created by LibreOffice during conversion
    extracted_images = {}
    if os.path.exists(output_dir):
        files_after = set(os.listdir(output_dir))
        new_files = files_after - files_before
        
        send(f"DEBUG: Files after conversion: {len(files_after)} files\n")
        send(f"DEBUG: New files created: {len(new_files)} files\n")
        
        if new_files:
            send(f"DEBUG: New files: {list(new_files)}\n")
        
        # Look for image files that match the LibreOffice naming pattern
        for filename in new_files:
            send(f"DEBUG: Checking file: {filename}\n")
            
            if (filename.startswith(output_basename + '_html_') and 
                filename.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.svg'))):
                
                image_path = os.path.join(output_dir, filename)
                send(f"DEBUG: Processing image: {image_path}\n")
                
                try:
                    with open(image_path, 'rb') as img_file:
                        image_data = img_file.read()
                        image_base64 = base64.b64encode(image_data).decode('utf-8')
                        
                        # Determine MIME type based on extension
                        ext = filename.lower().split('.')[-1]
                        mime_type = {
                            'png': 'image/png',
                            'jpg': 'image/jpeg', 
                            'jpeg': 'image/jpeg',
                            'gif': 'image/gif',
                            'svg': 'image/svg+xml'
                        }.get(ext, 'image/png')
                        
                        extracted_images[filename] = {
                            'data': image_base64,
                            'mime': mime_type
                        }
                        
                        send(f"DEBUG: Extracted image: {filename}, size: {len(image_data)} bytes\n")
                        
                    # Optionally remove the extracted file since we've captured it
                    # os.remove(image_path)
                    
                except Exception as e:
                    send(f"DEBUG: Error processing image {filename}: {str(e)}\n")

    document.dispose()
    document.close(True)
    
    # Send success response with extracted images data
    response = {'status': '200', 'images': extracted_images}
    send(json.dumps(response))


def listen():
    while True:
      message = sys.stdin.readline()
      if not message:
          break
      if message.rstrip() == "":
          send('204')
          continue
      try:
          fileOption = parser.parse_args(shlex.split(message.rstrip()))
      except (SystemExit, ValueError) as e:
          send('1')
          continue
      main()

if __name__ == "__main__":
    main()
