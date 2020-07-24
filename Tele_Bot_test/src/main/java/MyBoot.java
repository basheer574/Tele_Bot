import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.telegram.telegrambots.bots.TelegramLongPollingBot;
import org.telegram.telegrambots.meta.api.methods.GetFile;
import org.telegram.telegrambots.meta.api.methods.send.SendDocument;
import org.telegram.telegrambots.meta.api.methods.send.SendMessage;
import org.telegram.telegrambots.meta.api.objects.Update;
import org.telegram.telegrambots.meta.exceptions.TelegramApiException;
import java.io.*;

public class MyBoot extends TelegramLongPollingBot {

    public void onUpdateReceived(Update update) {
        if(update.getMessage().hasDocument()){
            File docFile = null;
            File pdfFile = new File("WordToPdf.pdf");
            GetFile file = new GetFile().setFileId(update.getMessage().getDocument().getFileId());
            try {
                SendMessage message = new SendMessage()
                        .setText("Converting...")
                        .setChatId(update.getMessage().getChatId());
                execute(message);
                File outputFile = new File(update.getMessage().getDocument().getFileName());
                String filePath = execute(file).getFilePath();
                docFile = downloadFile(filePath,outputFile);

                try{
                    ZipSecureFile.setMinInflateRatio(0);
                    InputStream stream = new FileInputStream(docFile);
                    XWPFDocument doc = new XWPFDocument(stream);
                    OutputStream out = new FileOutputStream(pdfFile);
                    PdfOptions options = PdfOptions.create();

                    PdfConverter.getInstance().convert(doc,out,options);
                    stream.close();
                    out.close();
                }catch (IOException e){
                    e.printStackTrace();
                }

            }catch (TelegramApiException e){
                e.printStackTrace();
            }
            SendDocument document = new SendDocument()
                    .setDocument(pdfFile)
                    .setChatId(update.getMessage().getChatId());
            try{
                execute(document);
            }catch (TelegramApiException e){
                e.printStackTrace();
            }
            docFile.delete();
            pdfFile.delete();
        }
    }
    

    public String getBotUsername() {
        return "PdfMatebot";
    }

    public String getBotToken() {
        return "1206144344:AAG4u9MlSmhCj_1jcXc1mfUgGAHQL6ZGC1M";
    }

}
