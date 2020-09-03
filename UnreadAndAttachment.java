/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package javamailtutorial;


import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;
import java.util.Random;
 
import javax.mail.Address;
import javax.mail.BodyPart;
import javax.mail.Flags;
import javax.mail.Folder;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.NoSuchProviderException;
import javax.mail.Part;
import javax.mail.Session;
import javax.mail.Store;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMultipart;
import javax.mail.search.FlagTerm;
 

public class UnreadAndAttachment {
    private String saveDirectory;
    private String bodyMessage = "";
    private int attnum = 1;

    public void setSaveDirectory(String dir) {
        this.saveDirectory = dir;
    }
 
    public void downloadEmailAttachments() throws Exception {
            
 //           String host = "outlook.office365.com";
  //          String username = "";
  //          String password = "";

            String host = "imap.gmail.com";
            String username = "";//to be inserted here
            String password = ""; //to be inserted here
            Properties props = new Properties();
            props.setProperty("mail.imap.ssl.enable", "true");
            
            Session session = Session.getInstance(props);

        try {
 
            // fetches new messages from server
            Store store = session.getStore("imap");
            store.connect(host, username, password);
              
            Folder inbox = store.getFolder("inbox");
            inbox.open(Folder.READ_ONLY);
            //System.out.println("count is:" + inbox.getUnreadMessageCount());
            Flags seen = new Flags(Flags.Flag.SEEN);
            FlagTerm unseenFlagTerm = new FlagTerm(seen,false);
            Message arrayMessages[] = inbox.search(unseenFlagTerm);
            System.out.println(arrayMessages.length);
            if(arrayMessages.length == 0) System.out.println("No new messages found");
            else{
              for (int i = 0; i < arrayMessages.length; i++) {
                System.out.println("**************************************************************");
                System.out.println(i);
                System.out.println("**************************************************************");
                Message message = arrayMessages[i];
                String contentType = message.getContentType();
                String attachFiles = "";
                dumpPart(arrayMessages[i]);
               // inbox.getMessage(message.getMessageNumber()).getContent();
                System.out.println("\t Message: " + bodyMessage);
                System.out.println("\t Attachments: " + attachFiles);
            }  
            }
            // disconnect
            inbox.close(false);
            store.close();
        } catch (NoSuchProviderException ex) {
            System.out.println("No provider for pop3.");
            ex.printStackTrace();
        } catch (MessagingException ex) {
            System.out.println("Could not connect to the message store");
            ex.printStackTrace();
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }
    
    private void dumpPart(Part p) throws Exception{
        //if(p instanceof Message)
            //dumpEnvelope((Message)p);
       String ct = p.getContentType();
       int level = 1;
       
       if(p.isMimeType("text/plain")){
          // System.out.println((String)p.getContent());
       }
       else if(p.isMimeType("multipart/*")){
           System.out.println("multipart");
           Multipart mp = (Multipart)p.getContent();
           level++;
           int count = mp.getCount();
           for(int i = 0; i < count; i++){
               dumpPart(mp.getBodyPart(i));
           }
           level--;
       }
       else if(p.isMimeType("message/rfc822")){
           level++;
           System.out.println("A nested message");
           dumpPart((Part)p.getContent());
           level--;
       }
       else{
           Object o = p.getContent();
           if(o instanceof String){
               bodyMessage = o.toString();
           //    System.out.println("Message is:" + (String)o);
           }
           
           else if(o instanceof InputStream){  
               if(!(level!=0 && !p.isMimeType("multipart/*"))){
                System.out.println("In here");
                   String originalFileName = p.getFileName();
                   String[] filenameArray = originalFileName.split("\\.",2);
                   String newFileName = filenameArray[0]+attnum+"."+filenameArray[1];
                   attnum++;
                   System.out.println("before else if");
                    ((MimeBodyPart)p).saveFile(saveDirectory + File.separator + newFileName);
                    System.out.println("after else if");
               }
                   
           }
           else{
                  System.out.println("unknown");
                bodyMessage = o.toString();
           System.out.println(o.toString());
           }
       }
       if(level!=0 && !p.isMimeType("multipart/*")){
           String disp = p.getDisposition();
           if(disp != null && disp.equalsIgnoreCase(Part.ATTACHMENT)){
              // String fileName = p.getFileName();
              
               System.out.println("before");
               String originalFileName = p.getFileName();
               System.out.println("original File Name " + originalFileName);
               String[] filenameArray = originalFileName.split("\\.",2);
               String newFileName = filenameArray[0]+attnum+"."+filenameArray[1];
               System.out.println("Newfilename " + newFileName);
               attnum = attnum + 1;
               ((MimeBodyPart)p).saveFile(saveDirectory + File.separator + newFileName);
               System.out.println("after");
               
           }
       }
       
    }
   
    public static void main(String[] args) throws Exception {
        
 
        String saveDirectory = "D:/Attachment";
 
        UnreadAndAttachment receiver = new UnreadAndAttachment();
        receiver.setSaveDirectory(saveDirectory);
        receiver.downloadEmailAttachments();
    }
}
