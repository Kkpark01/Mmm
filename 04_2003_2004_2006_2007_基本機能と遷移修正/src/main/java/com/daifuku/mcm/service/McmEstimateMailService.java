package com.daifuku.mcm.service;

import java.math.BigDecimal;
import java.util.*;
import jakarta.mail.internet.InternetAddress;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.mail.SimpleMailMessage;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import com.daifuku.mcm.common.McmConstants;

/** 申請と送信依頼を同時にコミット。SMTPはコミット後に実行し、失敗を再送可能な状態で残す。 */
@Service
public class McmEstimateMailService {
    private final JdbcTemplate jdbc;
    @Autowired(required=false) private JavaMailSender sender;
    @Value("${mcm.customer.mail-enabled:false}") private boolean enabled;
    public McmEstimateMailService(JdbcTemplate jdbc){this.jdbc=jdbc;}

    public String enqueue(BigDecimal estimate,String user){
        var users=jdbc.queryForList("SELECT LOGIN_ID,OYALOGIN_ID,MAILADDRESS,MITSUMORISHONIN_KIN FROM MCM.MCM_MO_TANTO WHERE LOGIN_ID=?",user);
        if(users.size()!=1)throw new IllegalStateException("申請者の担当者マスタを確認してください。");
        var source=users.get(0);
        BigDecimal amount=jdbc.queryForObject("SELECT TOP 1 MITSUMORI_GKIN FROM MCM.MCM_UM_MITSUMORI WHERE UM_KIHON_MITSUMORI_ID=? ORDER BY KAISI_DT,UM_MITSUMORI_ID",BigDecimal.class,estimate);
        // VB getShoninJohoの条件は申請者の承認限度額 > 最初の期間の見積金額。
        BigDecimal limit=source.get("MITSUMORISHONIN_KIN")==null?new BigDecimal("999999999999"):new BigDecimal(source.get("MITSUMORISHONIN_KIN").toString());
        String status="PENDING",from="",to="";
        if(limit.compareTo(amount==null?BigDecimal.ZERO:amount)<=0)status="SKIPPED";
        else {
            Object parent=source.get("OYALOGIN_ID");
            if(parent==null||parent.toString().isBlank()||parent.toString().equals(user))throw new IllegalStateException("申請先の上位担当者を担当者マスタに設定してください。");
            var targets=jdbc.queryForList("SELECT MAILADDRESS FROM MCM.MCM_MO_TANTO WHERE LOGIN_ID=?",parent);
            if(targets.size()!=1)throw new IllegalStateException("申請先の担当者マスタを確認してください。");
            from=address(source.get("MAILADDRESS"));to=address(targets.get(0).get("MAILADDRESS"));
            configured();
        }
        String key=UUID.randomUUID().toString();
        jdbc.update("INSERT INTO MCM.MCM_WEB_MAIL_OUTBOX (MAIL_ID,UM_KIHON_MITSUMORI_ID,REQUESTED_BY,FROM_ADDRESS,TO_ADDRESS,SUBJECT,BODY,STATUS,ATTEMPTS,CREATED_DT,UPDATED_DT) VALUES (?,?,?,?,?,?,?,?,0,GETDATE(),GETDATE())",key,estimate,user,from,to,McmConstants.TENPO_SHINSA_TITLE,McmConstants.TENPO_SHINSA_BODY,status);
        return key;
    }
    private String address(Object raw){
        try{if(raw==null)throw new IllegalArgumentException();String value=raw.toString().trim();if(value.contains("\r")||value.contains("\n"))throw new IllegalArgumentException();var a=InternetAddress.parse(value,true);if(a.length!=1)throw new IllegalArgumentException();a[0].validate();return a[0].getAddress();}
        catch(Exception ex){throw new IllegalStateException("申請者・申請先のメールアドレスを担当者マスタに設定してください。");}
    }
    private void configured(){if(!enabled||sender==null)throw new IllegalStateException("申請メールのSMTP接続が未設定です。メール設定を完了してから申請してください。");}

    /** HTTP更新処理のトランザクション完了後に呼び出す。二重リクエストは状態更新で排除。 */
    public String dispatchLatest(BigDecimal estimate,String user,boolean retry){
        var rows=jdbc.queryForList("SELECT TOP 1 * FROM MCM.MCM_WEB_MAIL_OUTBOX WHERE UM_KIHON_MITSUMORI_ID=? AND REQUESTED_BY=? ORDER BY CREATED_DT DESC,MAIL_ID",estimate,user);
        if(rows.isEmpty())return "NONE";
        var row=rows.get(0);String state=String.valueOf(row.get("STATUS"));
        if(!state.equals("PENDING")&&!(retry&&state.equals("FAILED")))return state;
        configured();
        String key=String.valueOf(row.get("MAIL_ID"));
        if(jdbc.update("UPDATE MCM.MCM_WEB_MAIL_OUTBOX SET STATUS='SENDING',ATTEMPTS=ATTEMPTS+1,UPDATED_DT=GETDATE() WHERE MAIL_ID=? AND STATUS=?",key,state)!=1)return "SENDING";
        var message=new SimpleMailMessage();message.setFrom(String.valueOf(row.get("FROM_ADDRESS")));message.setTo(String.valueOf(row.get("TO_ADDRESS")));message.setSubject(String.valueOf(row.get("SUBJECT")));message.setText(String.valueOf(row.get("BODY")));
        try{sender.send(message);}catch(org.springframework.mail.MailException ex){jdbc.update("UPDATE MCM.MCM_WEB_MAIL_OUTBOX SET STATUS='FAILED',LAST_ERROR=?,UPDATED_DT=GETDATE() WHERE MAIL_ID=?",ex.getClass().getSimpleName(),key);return "FAILED";}
        // SMTP受理後のDBエラーはSENDINGを維持し、自動再送しない（重複防止）。
        jdbc.update("UPDATE MCM.MCM_WEB_MAIL_OUTBOX SET STATUS='SENT',LAST_ERROR=NULL,UPDATED_DT=GETDATE() WHERE MAIL_ID=?",key);return "SENT";
    }
    public String status(BigDecimal estimate,String user){
        var rows=jdbc.queryForList("SELECT TOP 1 STATUS FROM MCM.MCM_WEB_MAIL_OUTBOX WHERE UM_KIHON_MITSUMORI_ID=? AND REQUESTED_BY=? ORDER BY CREATED_DT DESC,MAIL_ID",estimate,user);
        return rows.isEmpty()?"NONE":String.valueOf(rows.get(0).get("STATUS"));
    }
    public static String message(String status){return switch(status){case "SENT"->"申請を完了し、審査依頼メールを送信しました。";case "FAILED"->"申請は完了しましたが、メールを送信できませんでした。送信状況を確認し、メール再送を実行してください。";case "SENDING"->"申請は完了しました。メールの送信結果を確認中です。管理者に送信状況を確認してください。";case "SKIPPED"->"申請を完了しました。VB版の金額条件によりメール通知の対象外です。";default->"申請を受け付けました。審査依頼メールは送信待ちです。";};}
}
