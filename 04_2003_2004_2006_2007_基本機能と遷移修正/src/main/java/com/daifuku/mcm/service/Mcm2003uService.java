package com.daifuku.mcm.service;

import java.math.BigDecimal;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import com.daifuku.mcm.common.McmConstants;
import com.daifuku.mcm.constants.Mcm2003uConstants;
import com.daifuku.mcm.form.Mcm2003uForm;
import com.daifuku.mcm.repository.Mcm2003uRepository;

/**
 * 【変換元】Mcm2003uScreen.vb（MitsumoriButton_Click / HakkoButton_Click / SinseiButton_Click）
 * MCM2003U 店舗見積内容基本設定サービス
 *
 * 保存・集計・発行・申請のDB更新は同じトランザクション。SMTPはコミット後に送信する。
 */
@Service
public class Mcm2003uService {
    private static final String UNLOCK="mcm2003u.adminUnlock";
    /** HTTP入力にはバインドしない、見積・利用者・遷移元に限定したセッション状態。 */
    private record Unlock(java.math.BigDecimal id,String user,int entry) implements java.io.Serializable {}

    public boolean canReleaseLock(jakarta.servlet.http.HttpSession session) {
        var auth=org.springframework.security.core.context.SecurityContextHolder.getContext().getAuthentication();
        if(auth==null||!auth.isAuthenticated()||auth.getAuthorities().stream().noneMatch(a->a.getAuthority().equals(com.daifuku.mcm.common.AppConstants.ROLE_UPDATE)))return false;
        Object division=session.getAttribute(com.daifuku.mcm.common.AppConstants.SESSION_AUTHORITY_DIVISION);
        if("0".equals(division)||"1".equals(division))return false;
        try{return repo.canReleaseLock(auth.getName());}catch(org.springframework.dao.DataAccessException ex){return false;}
    }
    public void clearUnlock(jakarta.servlet.http.HttpSession session){session.removeAttribute(UNLOCK);}
    public boolean isUnlocked(BigDecimal id,jakarta.servlet.http.HttpSession session) {
        if(!(session.getAttribute(UNLOCK) instanceof Unlock grant))return false;
        var auth=org.springframework.security.core.context.SecurityContextHolder.getContext().getAuthentication();
        Object currentId=session.getAttribute("mcm2003u.umKihonMitsumoriId"),entry=session.getAttribute("mcm2003u.seniMotoKbn");
        boolean valid=auth!=null&&grant.user().equals(auth.getName())&&com.daifuku.mcm.common.CustomerScreenSupport.same(id,grant.id())
            &&currentId instanceof BigDecimal current&&com.daifuku.mcm.common.CustomerScreenSupport.same(current,id)
            &&grant.entry()==(entry instanceof Integer mode?mode:3)&&session.getAttribute("mcm2003u.draft")==null&&canReleaseLock(session);
        if(!valid)clearUnlock(session);return valid;
    }
    public Mcm2003uForm releaseLock(BigDecimal id,int entry,jakarta.servlet.http.HttpSession session){
        clearUnlock(session);
        if(!canReleaseLock(session)||id==null||session.getAttribute("mcm2003u.draft")!=null)throw new IllegalStateException("管理者用ロック解除の権限がありません。");
        if(!(session.getAttribute("mcm2003u.umKihonMitsumoriId") instanceof BigDecimal current)||!com.daifuku.mcm.common.CustomerScreenSupport.same(current,id)
            ||entry!=(session.getAttribute("mcm2003u.seniMotoKbn") instanceof Integer mode?mode:3))throw new IllegalStateException("画面の有効期限が切れています。再検索してください。");
        var latest=load(id);latest.setSeniMotoKbn(entry);
        if(latest.getUmKihonMitsumoriId()==null)throw new IllegalStateException("見積情報を確認してください。");
        var auth=org.springframework.security.core.context.SecurityContextHolder.getContext().getAuthentication();
        session.setAttribute(UNLOCK,new Unlock(id,auth.getName(),entry));return latest;
    }
    private boolean unlockedForThisOperation(BigDecimal id){
        var attributes=org.springframework.web.context.request.RequestContextHolder.getRequestAttributes();
        if(!(attributes instanceof org.springframework.web.context.request.ServletRequestAttributes web))return false;
        var request=web.getRequest();var session=request.getSession(false);if(session==null)return false;
        String route=request.getRequestURI().substring(request.getContextPath().length());
        // 2005Uや再選定の保存には解除を波及させない。
        return "POST".equals(request.getMethod())&&java.util.Set.of("/mcm2003u/save","/mcm2003u/publish","/mcm2003u/attachment/add","/mcm2003u/attachment/remove").contains(route)&&isUnlocked(id,session);
    }

    @Autowired
    private Mcm2003uRepository repo;
    @Autowired private McmCustomerTotalsService totals;
    @Autowired private McmCustomerIntegrityService integrity;
    @Autowired private McmCustomerReportService reports;
    @Autowired private McmEstimateMailService mail;

    // ===================================================================
    // ロード
    // 【変換元】Mcm2003uScreen.vb — SearchRead()
    // ===================================================================

    public Mcm2003uForm load(BigDecimal umKihonMitsumoriId) {
        Mcm2003uForm form = repo.findKihonMitsumori(umKihonMitsumoriId);
        form.setMitsumoriRows(repo.findMitsumoriRows(umKihonMitsumoriId));
        form.setTenpuRows(repo.findTenpuRows(umKihonMitsumoriId));
        form.setBrandRows(repo.findBrandRows(umKihonMitsumoriId));
        form.setKoseiRows(repo.findKoseiRows(umKihonMitsumoriId));
        form.setMeisaiRows(repo.findMeisaiRows(umKihonMitsumoriId));
        return form;
    }

    // ===================================================================
    // 見積登録（保存）
    // 【変換元】MitsumoriButton_Click → UpdateButton(BUTTON_FLG_SAKUSEI) → SenteiNashiUpdate
    // ===================================================================

    @Transactional
    public void save(Mcm2003uForm form,String user){requireEditable(form);repo.updateBrandMethods(form,user);repo.updateKihonMitsumori(form,user);integrity.estimate(form.getUmKihonMitsumoriId());totals.recalculate(form.getUmKihonMitsumoriId(),user);}

    // ===================================================================
    // 見積発行
    // 【変換元】HakkoButton_Click → UpdateButton(BUTTON_FLG_HAKKO) → SenteiNashiUpdate
    //   状態が作成中(9)の場合のみ見積(1)に変更。
    //   帳票生成が失敗した場合、保存・状態更新もロールバックする。
    // ===================================================================

    @Transactional
    public byte[] publish(Mcm2003uForm form,String user){
        requireEditable(form);repo.updateBrandMethods(form,user);repo.updateKihonMitsumori(form,user);integrity.estimate(form.getUmKihonMitsumoriId());totals.recalculate(form.getUmKihonMitsumoriId(),user);
        if(McmConstants.JOTAI_SAKUSEICHU.equals(form.getJotai()))repo.updateJotai(form.getUmKihonMitsumoriId(),McmConstants.JOTAI_MITSUMORI,user);
        return reports.generate(form.getUmKihonMitsumoriId());
    }

    // ===================================================================
    // 申請
    // 【変換元】SinseiButton_Click
    //   添付ファイル必須チェック → SYOUNIN_JOTAI を審査中(1)へ変更
    //   同一トランザクション内でメール送信依頼を保存する。
    // ===================================================================

    @Transactional
    public void saveAndApply(Mcm2003uForm form,String user){
        requireEditable(form);
        if(repo.countTenpu(form.getUmKihonMitsumoriId())==0)throw new IllegalStateException("見積資料を添付してください。");
        repo.updateBrandMethods(form,user);repo.updateKihonMitsumori(form,user);integrity.estimate(form.getUmKihonMitsumoriId());totals.recalculate(form.getUmKihonMitsumoriId(),user);
        mail.enqueue(form.getUmKihonMitsumoriId(),user);
        repo.updateSyouninJotai(form.getUmKihonMitsumoriId(),McmConstants.SHONINJOTAI_SHINSACHU_CD,user);
    }

    /** 入力の保存と申請を同じトランザクションで行い、未保存内容の欠落を防ぐ。 */
    @Transactional
    public void sinsei(BigDecimal id, String user) { saveAndApply(repo.findKihonMitsumori(id),user); }
    public void requireEditable(Mcm2003uForm form) {
        if (form.getUmKihonMitsumoriId()==null || form.getUmKihonMitsumoriId().signum()<=0) throw new IllegalStateException("見積を選択してください。");
        repo.lock(form.getUmKihonMitsumoriId());
        var current=repo.findKihonMitsumori(form.getUmKihonMitsumoriId());
        if (!unlockedForThisOperation(form.getUmKihonMitsumoriId()) && (isReadOnly(current.getSyouninJotai(),form.getSeniMotoKbn()) || "3".equals(current.getJotai()) || "4".equals(current.getJotai())))
            throw new IllegalStateException("見積の状態が変更されています。再読み込みしてください。");
        if (!java.util.Objects.equals(current.getLastupdateDt(),form.getLastupdateDt())) throw new IllegalStateException("他のユーザによって更新されています。再読み込みしてください。");
        repo.prepareHeaderChanges(form,current);
        if (form.getMitsumoriLevel()!=null && !form.getMitsumoriLevel().isBlank() && !form.getMitsumoriLevel().matches("[0-9]{1,2}"))
            throw new IllegalStateException("見積レベルは数値で入力してください。");
    }

    // ===================================================================
    // ReadOnly判定
    // 【変換元】View() — 承認状態が審査中/承認中/承認済、または遷移元区分=4の場合ReadOnly
    // ===================================================================

    public boolean isReadOnly(String syouninJotai, int seniMotoKbn) {
        if (seniMotoKbn == Mcm2003uConstants.SENIMOTO_KBN_SHONIN) {
            return true;
        }
        return McmConstants.SHONINJOTAI_SHINSACHU_CD.equals(syouninJotai)
            || McmConstants.SHONINJOTAI_SHONINCHU_CD.equals(syouninJotai)
            || McmConstants.SHONINJOTAI_SHONINZUMI_CD.equals(syouninJotai);
    }
}
