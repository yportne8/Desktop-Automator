import traceback
from typing import Union
from win32com import client


class Outlook:
    """
    _A light wrapper around Outlook for desktop._
    """

    def __init__(self):
        pass
        
    @property
    def app(self):
        return client.Dispatch("Outlook.Application")

    def __draft(self, body, subject, to, attachments, cc, mark_important):
        """_Hidden method for draft_email_

        Args:
            body (_type_): _email body, html or text_
            subject (_type_): _subject line_
            to (_type_): _to email address(es)_
            cc (_type_): _cc email address(es)_
            attachments (_type_): _description_

        Returns:
            _type_: _description_
        """
        app = self.app
        mail = app.createItem(0)
        
        if subject:
            mail.Subject = subject
        
        # [NOTE] Htmlbody is used by outlook if both are assigned.
        # Body assigned html string as backup in case (markup???).
        # [TODO] shortcut for html string determination.
        if "<" in body and ">" in body:  
            mail.HTMLBody = body
            mail.body = body
        else:
            mail.Body = body
            
        if to:
            if type(to) == list:
                sendTo = to[0]
                for address in to[1:]:
                    sendTo += f"; {address}"
                to = sendTo
            mail.To = to
        
        if attachments:    
            for attachment in attachments: mail.Attachments.Add(attachment)

        if cc:
            if type(cc) == list:
                sendCc = cc[0]
                for address in cc[1:]:
                    sendCc += f"; {address}"
                cc = sendCc
            mail.Cc = cc

        if mark_important: mail.Importance = 2 
        return mail
    
    def send_email(self, body: str,
                   subject: str=" ", to: Union[str, list]=None,
                   attachments: list=None, cc: Union[str, list]=None,
                   mark_important: bool=False):
        """_Sends an email in the background according to specifications._

        Args:
            to (Union[str, list], optional): _to email address(es)_. Defaults to None.
            cc (Union[str, list], optional): _cc email address(es)_. Defaults to None.
            body (str, optional): _email body, html or text_. Defaults to " ".
            subject (str, optional): _subject line_. Defaults to " ".
            attachments (list, optional): _list of c:\\file\\paths.xlsx_. Defaults to None.
        """

        # [TODO] Using " " to bypass any content-less email errors
        # raised by Outlook (if they exist || should be bypassed...)
        mail = self.__draft(body,subject,to,attachments,cc,mark_important)
        try:
            mail.Send()
        
        except Exception as e:
            traceback.print_exception(e)
    
    def draft_email(self, body: str,
                   subject: str=" ", to: Union[str, list]=None,
                   attachments: list=None, cc: Union[str, list]=None,
                   close: bool=True, mark_important: bool=False):
        """_Drafts an email per specifications. Drafted emails are
            always saved to the drafts folder of Outlook._

        Args:
            body (str): _description_
            subject (str, optional): _description_. Defaults to " ".
            to (Union[str, list], optional): _description_. Defaults to None.
            attachments (list, optional): _description_. Defaults to None.
            cc (Union[str, list], optional): _description_. Defaults to None.
            close (bool, optional): _description_. Defaults to False.
        """

        __ignore=False
        mail = self.__draft(body,subject,to,attachments,cc,mark_important)
        mail.Close(__ignore) if close else mail.Display()