import java.io.*;
import java.net.*;
import java.util.*;
import javax.servlet.*;
import javax.servlet.http.*;

/**
 * SendMailServlet is used to send mail from a web server process using SMTP.
 * It reads the following parameters from the HTTP request:
 *
 * <pre>
 * host        - Name or IP address of mail host.
 * domain      - Domain of sender.
 * sender      - Address of sender.
 * recipients  - Addresses of recipients (separated with commas).
 * subject     - Subject of e-mail.
 * mailtext    - Text of e-mail message.
 * </pre>
 *
 * <b>Notes:</b>
 * <ul>
 *    <li>The parameter names are case-senstive and must be specified in the
 *        HTML document exactly as they are shown above.
 *
 *    <li>The SMTP port is assumed to be 25.
 * </ul>
 *
 * @author  Thornton Rose
 * @version 1.0
 */
public class SendMailServlet extends HttpServlet
{
   // Constants

   private static final int  SMTP_PORT = 25;
   private static final char SMTP_ERROR_CODE1 = '4';
   private static final char SMTP_ERROR_CODE2 = '5';

   // Methods
   
   /**
    * getServletInfo() returns the servlet description.
    *
    * @return The servlet description.
    */
   public String getServletInfo() {
      return "SendMailServlet 1.0";
   }

   /**
    * service() is used to service the HTTP request. It handles the HEAD, GET,
    * and POST HTTP methods.
    *
    * @param request  The HTTP request object.
    * @param response The HTTP response object.
    *
    * @exception ServletException If a ServletException occurs, it is passed up
    *    the call chain; no special handling is done.
    *
    * @exception IOException If an IO exception occurs while not connected to 
    *    the SMTP server, it is passed up the call chain. Otherwise, the
    *    IOException is caught and a response is sent to the client.
    */
   public void service(HttpServletRequest request, HttpServletResponse response)
      throws ServletException, IOException {

      String host;
      String domain;
      String sender;
      String recipients;
      String subject;
      String mailtext;
      String maildata;
      Vector sessionTrace = new Vector(20);

      // Get the HTTP parameters.
      
      host = getParameter(request, "host");
      domain = getParameter(request, "domain");
      sender = getParameter(request, "sender");
      recipients = getParameter(request, "recipients");
      subject = getParameter(request, "subject");
      mailtext = getParameter(request, "mailtext");
      
      // Try to send the mail, then the response, catching
      // any IO exception.

      try {
         // Send the mail.

         maildata =
            "Date: " + (new Date()).toString() + "\r\n" +
            "From: " + sender + "\r\n" +
            "To: " + recipients + "\r\n" +
            "Subject: " + subject + "\r\n" +
            "\r\n" +
            mailtext + "\r\n";
         sendMail(host, domain, sender, recipients, subject, maildata, sessionTrace);

         // Send a response page to the client.

         sendResponse(
            request, 
            response, 
            "The mail was queued for delivery.", 
            sessionTrace);
      }
      catch(IOException theException) {
         // Send a response page to the client.

         sendResponse(request, response, theException.toString(), sessionTrace);
      }
   }

   /**
    * getParameter() gets the value of the named parameter from the HTTP 
    * request.
    *
    * @param name The name of the parameter to get.
    *
    * @return The value of the named parameter.
    */
   protected String getParameter(HttpServletRequest request, String name) {
      String[] values;
      String result = "";

      values = request.getParameterValues(name);

      if (values != null) {
         result = values[0];
      }

      return result;
   }

   /**
    * sendMail() sends the mail.
    *
    * @param host          Host name or IP.
    * @param domain        Sender's domain.
    * @param sender        Sender's address.
    * @param recipients    Recipient addresses.
    * @param subject       Subject line.
    * @param maildata      Mail data.
    * @param sessionTrace  List or commands send and replies read.
    *
    * @exception IOException
    */
   protected void sendMail(
      String host,
      String domain,
      String sender,
      String recipients,
      String subject,
      String maildata,
      Vector sessionTrace) throws IOException {

      Socket            mailSocket;
      BufferedReader    socketIn;
      DataOutputStream  socketOut;
      String            address;
      StringTokenizer   tokenizer;

      // Open the connection to the SMTP server, then get references to the
      // input and output streams.

      mailSocket = new Socket(host, SMTP_PORT);
      socketIn = new BufferedReader(
         new InputStreamReader(mailSocket.getInputStream()) );
      socketOut = new DataOutputStream(mailSocket.getOutputStream());

      // Get the initial reply from the server.

      readReply(socketIn, sessionTrace);

      // Greet the server.

      sendCommand(socketOut, "HELO " + domain, sessionTrace);
      readReply(socketIn, sessionTrace);

      // Send the sender's address.
      
      sendCommand(socketOut, "MAIL FROM: " + sender, sessionTrace);
      readReply(socketIn, sessionTrace);

      // Send the list of recipients.

      tokenizer = new StringTokenizer(recipients, ",");

      while (tokenizer.hasMoreElements()) {
         sendCommand(socketOut, "RCPT TO: " + tokenizer.nextToken(), sessionTrace);
         readReply(socketIn, sessionTrace);
      }

      // Start the data section.

      sendCommand(socketOut, "DATA", sessionTrace);
      readReply(socketIn, sessionTrace);

      // Send the mail message.
      
      sendCommand(socketOut, maildata + ".", sessionTrace);
      readReply(socketIn, sessionTrace);

      // End the session.
      
      sendCommand(socketOut, "QUIT", sessionTrace);
      readReply(socketIn, sessionTrace);
   }

   /**
    * sendCommand() sends an SMTP command to the SMTP server. An
    * SMTP command is a string that look like:
    * <pre>
    * [key words] [data][CR][LF]
    * </pre>
    *
    * <p>Example: <pre>HELO xyz.com</pre>
    *
    * @param out          The output stream to which to write the command.
    * @param command      The command to write.
    * @param sessionTrace List of commands sent and replies read.
    *
    * @exception IOException No special handling is done; the exception is
    *    passed back up the call chain.
    */
   private void sendCommand(DataOutputStream out, String command, Vector sessionTrace) 
      throws IOException {
      
      out.writeBytes(command + "\r\n");
      sessionTrace.addElement(command);
      // System.out.println(command);
   }

   /**
    * readReply() reads the reply from the SMTP server.
    *
    * @param reader       The input reader from which to read the reply.
    * @param sessionTrace List of commands sent and replies read.
    *
    * @exception IOException If an IOException occurs because of an error
    *    with socket IO, no special handling is done, and the exception is
    *    passed back up the call chain. If the status code in the reply from
    *    the SMTP server is equal to SMTP_ERROR_CODE1 or SMTP_ERROR_CODE2, an
    *    IOException containing "SMTP: " + the reply is thrown.
    */
   private void readReply(BufferedReader reader, Vector sessionTrace)
      throws IOException {
      
      String reply;
      char   statusCode;
      
      reply = reader.readLine();
      statusCode = reply.charAt(0);
      sessionTrace.addElement(reply);
      // System.out.println(reply);

      if ( (statusCode == SMTP_ERROR_CODE1) | 
           (statusCode == SMTP_ERROR_CODE2) ) {
         throw (new IOException("SMTP: " + reply));
      }
   }
   
   /**
    * sendResponse() sends the response to the client that generated the HTTP 
    * request.
    *
    * @param request       The HTTP response object.
    * @param response      The HTTP response object.
    * @param statusMessage The status message.
    * @param sessionTrace  List of commands sent and replies read.
    * 
    * @exception IOException An IOException is thrown if an error occurs
    *    while writing to the output stream of the HTTP response.
    */
   protected void sendResponse(
      HttpServletRequest request, 
      HttpServletResponse response, 
      String statusMessage, 
      Vector sessionTrace) throws IOException {
      
      ServletOutputStream out;

      // Get a reference to the HTTP response writer.

      response.setContentType("text/html");
      out = response.getOutputStream();
      
      // Send the header.

      out.println("<html>");
      out.println("<head>");
      out.println("<title>" + getServletInfo() + "</title>");
      out.println("</head>");
      out.println("<body>");
      out.println("<h2>" + getServletInfo() + "</h2>");
      
      // Send the status message.
      
      out.println("<p>" + statusMessage);
      
      // Send the request parameters.
      
      out.println("<p><b>Request Parameters:</b>");
      out.println("<pre>");
      out.println("host       = " + getParameter(request, "host"));
      out.println("domain     = " + getParameter(request, "domain"));
      out.println("sender     = " + getParameter(request, "sender"));
      out.println("recipients = " + getParameter(request, "recipients"));
      out.println("subject    = " + getParameter(request, "subject"));
      out.println("mailtext   = " + getParameter(request, "mailtext"));
      out.println("</pre>");
      
      // Send the session trace.
      
      out.println("<p><b>Session Trace:</b>");
      out.println("<pre>");
      
      Enumeration e = sessionTrace.elements();
      
      while (e.hasMoreElements()) {
         out.println((String) e.nextElement());
      }
      
      out.println("</pre>");

      // Send the footer.
      
      out.println("</body>");
      out.println("</html>");
   }
}

