package com.daifuku.mcm.service;

import com.daifuku.mcm.common.McmConstants;
import com.daifuku.mcm.exception.McmBusinessException;
import com.daifuku.mcm.form.Mcm1011uForm.KeiyakuRowForm;
import java.math.BigDecimal;
import java.util.*;
import jakarta.mail.internet.InternetAddress;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.mail.SimpleMailMessage;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.stereotype.Service;
import org.springframework.transaction.support.*;

/** VB setMailJoho/getShoninJoho: 審査完了時のみ、直属の上位担当者へ重複を除いて通知。 */
@Service
public class Mcm1011uMailService {
    @Autowired private JdbcTemplate jdbc;
    @Autowired(required=false) private JavaMailSender sender;
    private static final org.slf4j.Logger log=org.slf4j.LoggerFactory.getLogger(Mcm1011uMailService.class);

    public List<SimpleMailMessage> prepare(List<KeiyakuRowForm> rows, String user) {
        if (rows.stream().noneMatch(r -> "1".equals(r.getContractState()))) return List.of();
        var source=jdbc.queryForList("SELECT LOGIN_ID, OYALOGIN_ID, MAILADDRESS, KEIYAKUSHONIN_KIN FROM MCM.MCM_MO_TANTO WHERE LOGIN_ID=?",user);
        if(source.size()!=1)throw new McmBusinessException("担当者マスタの申請者情報を確認してください。");
        var from=source.get(0);
        BigDecimal limit=from.get("KEIYAKUSHONIN_KIN")==null ? new BigDecimal("999999999999") : new BigDecimal(from.get("KEIYAKUSHONIN_KIN").toString());
        // 原VBは申請者の承認限度額 > 選択期間の仕切合計の場合に通知する。
        if(rows.stream().noneMatch(r -> "1".equals(r.getContractState()) && limit.compareTo(r.getKingaku())>0))return List.of();
        Object parent=from.get("OYALOGIN_ID");
        if(parent==null || parent.toString().isBlank())return List.of(); // VB: 上位担当者なしは通知なし。
        if(parent.toString().equals(user))throw new McmBusinessException("担当者マスタの上位担当者が本人になっています。");
        var targets=jdbc.queryForList("SELECT MAILADDRESS FROM MCM.MCM_MO_TANTO WHERE LOGIN_ID=?",parent);
        if(targets.isEmpty())return List.of();
        if(targets.size()!=1)throw new McmBusinessException("担当者マスタの上位担当者情報を確認してください。");
        Object rawFrom=from.get("MAILADDRESS"),rawTo=targets.get(0).get("MAILADDRESS");
        if(rawFrom==null || rawTo==null || rawFrom.toString().isBlank() || rawTo.toString().isBlank())return List.of();
        if(sender==null)throw new McmBusinessException("承認依頼メールの送信設定がありません。管理者に確認してください。");
        var message=new SimpleMailMessage();message.setFrom(address(rawFrom));message.setTo(address(rawTo));
        message.setSubject(McmConstants.TORI_SHONIN_TITLE);message.setText(McmConstants.TORI_SHONIN_BODY);
        return List.of(message); // 同一操作・同一ログイン利用者の宛先は1件。
    }
    private String address(Object value) {
        try {
            String text=value.toString().trim();
            if(text.contains("\r")||text.contains("\n"))throw new IllegalArgumentException();
            var addresses=InternetAddress.parse(text,true);
            if(addresses.length!=1)throw new IllegalArgumentException();
            addresses[0].validate();return addresses[0].getAddress();
        } catch(Exception ex){throw new McmBusinessException("担当者マスタのメールアドレスを確認してください。");}
    }
    /** DBコミット後に通知。失敗を承認失敗と誤表示せず、再承認による二重処理を防ぐ。 */
    public void afterCommit(List<SimpleMailMessage> messages, List<String> warnings) {
        if(!TransactionSynchronizationManager.isSynchronizationActive())throw new IllegalStateException("Mail requires a transaction");
        TransactionSynchronizationManager.registerSynchronization(new TransactionSynchronization(){
            @Override public void afterCommit(){
                for(var message:messages)try{sender.send(message);}catch(Exception ex){
                    log.error("MCM1011U 承認済・通知失敗。管理者による通知確認が必要",ex);
                    warnings.add("審査・承認は完了しましたが、承認依頼メールを送信できませんでした。管理者に通知状況を確認してください。");
                }
            }
        });
    }
}
