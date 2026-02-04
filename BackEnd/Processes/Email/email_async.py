import threading

from BackEnd.Processes.Email.email_service import send_email


def send_login_email_async(email, project_name, work_order):
    
    subject = f"New login created -WO {work_order}"
    body = f"""
            <html>

<body style="margin:0; padding:0; background-color:#f4f6f8;">
    <table width="100%" cellpadding="0" cellspacing="0">
        <tr>
            <td style="padding:10px 20px; background-color:#5ea3cc;">
                <table width="100%" cellpadding="0" cellspacing="0">
                    <tr>
                        <td width="160" align="left" style="padding-left:10px;">
                            <img src="https://github.com/juliandevatbs/EMAIL-LOGO/blob/main/LOGO_SRL_FINAL.png?raw=true"
                                width="100" alt="SRL Logo" style="display:block;">
                        </td>
                        <td align="left" style="color:#ffffff; font-size:18px; font-weight:bold; padding-left:10px;">
                            Southern Research Laboratories, Inc.
                        </td>
                    </tr>
                </table>
            </td>
        </tr>


        <tr>
            <td style="padding:20px 25px 25px 25px;">
                <table width="100%" cellpadding="8" cellspacing="0" style="border-collapse:collapse;">
                    <tr style="background-color:#f1f3f5;">
                        <td style="border:1px solid #dddddd;"><strong>Project Name</strong></td>
                        <td style="border:1px solid #dddddd;">{project_name}</td>
                    </tr>
                    <tr>
                        <td style="border:1px solid #dddddd;"><strong>Work Order</strong></td>
                        <td style="border:1px solid #dddddd;">{work_order}</td>
                    </tr>
                    <tr style="background-color:#f1f3f5;">
                        <td style="border:1px solid #dddddd;"><strong>Receipt Date</strong></td>
                        <td style="border:1px solid #dddddd;"></td>
                    </tr>
                    <tr>
                        <td style="border:1px solid #dddddd;"><strong>Estimated Delivery Date</strong></td>
                        <td style="border:1px solid #dddddd;"></td>
                    </tr>
                </table>
            </td>
        </tr>


        <tr>
            <td style="padding:0 25px 25px 25px;">
                <table width="100%" cellpadding="8" cellspacing="0" style="border-collapse:collapse;">
                    <tr style="background-color:#0ca4cf; color:#ffffff;">
                        <td style="border:1px solid #0ca4cf;"><strong>Client Sample ID</strong></td>
                        <td style="border:1px solid #0ca4cf;"><strong>Matrix</strong></td>
                        <td style="border:1px solid #0ca4cf;"><strong>Analysis</strong></td>
                        <td style="border:1px solid #0ca4cf;"><strong>Quantity</strong></td>
                    </tr>

 
                </table>
            </td>
        </tr>

        <tr>
            <td style="padding:20px 25px;">
                <span style="font-size:14px;">
                    If you have any questions regarding your samples or analysis, please contact us                </span> 
            </td>
        </tr>

        <tr>
            <td style="padding:15px; background-color:#f1f3f5; text-align:center; font-size:12px; color:#777777;">
                © 2026 SRL<br>
                This is an automated message. Please do not reply.
            </td>
        </tr>

    </table>

    </td>
    </tr>
    </table>
</body>

</html>
    
    """
    
    
    threading.Thread(
        
        target = send_email,
        args=(email, subject, body),
        daemon=True
        
        
    ).start()