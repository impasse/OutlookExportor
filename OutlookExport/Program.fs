open System.IO
open System.Runtime.InteropServices
open Microsoft.Office.Interop.Outlook

let initDirectory () =
    if Directory.Exists("HTML") |> not then
        Directory.CreateDirectory("HTML") |> ignore
    else
        for file in Directory.EnumerateFiles("HTML") do
            File.Delete(file)
    if Directory.Exists("Attachments") |> not then
        Directory.CreateDirectory("Attachments") |> ignore
    else
        for file in Directory.EnumerateFiles("Attachments") do
            File.Delete(file)

let safeFilename (filename:string) = 
    filename.Split(Path.GetInvalidFileNameChars()) |> String.concat "_"

let saveMailItem (item:MailItem) =
    printfn "Processing %A" item.Subject
    let filename = sprintf "%d-%d-%d--%s.html" item.ReceivedTime.Year item.ReceivedTime.Month item.ReceivedTime.Day item.Subject
    let path = Path.Combine("HTML", safeFilename filename)
    File.WriteAllText(path, item.HTMLBody) |> ignore
    for attachment in item.Attachments do
        printfn "Processing attachment %A" attachment.FileName
        let savePath = Path.GetFullPath(Path.Combine("Attachments", safeFilename(attachment.FileName)))
        attachment.SaveAsFile(savePath)
 

let rec iteratee (folder:MAPIFolder) =
    if folder.DefaultItemType = OlItemType.olMailItem then
        for subFolder in folder.Folders do
            iteratee subFolder
        printfn "Start procession folder: %A" folder.Name
        for item in folder.Items do
            match item with
                | :? MailItem as mi -> saveMailItem mi
                | _ -> ()
            Marshal.ReleaseComObject(item) |> ignore
        Marshal.ReleaseComObject(folder) |> ignore

[<EntryPoint>]
let main argv = 
    initDirectory ()
    let app = new ApplicationClass()
    let mapi = app.GetNamespace("MAPI")
    let root = mapi.DefaultStore.GetRootFolder()
    iteratee root
    printfn "Finish"
    0
