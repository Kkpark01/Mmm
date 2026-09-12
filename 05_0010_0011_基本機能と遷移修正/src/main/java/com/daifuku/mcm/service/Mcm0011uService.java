package com.daifuku.mcm.service;

import java.math.BigDecimal;
import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import java.time.LocalDateTime;
import java.util.HexFormat;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import java.util.stream.Collectors;

import jakarta.validation.Validator;
import org.springframework.beans.BeanUtils;
import org.springframework.beans.BeanWrapperImpl;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.jdbc.core.BeanPropertyRowMapper;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Isolation;
import org.springframework.transaction.annotation.Transactional;
import com.daifuku.mcm.entity.NonyusakiEntity;
import com.daifuku.mcm.form.Mcm0011uForm;

/** VB MCM0011Uの納入先保存・削除。共有Entity/Repositoryを変更せず、VBの実列を使用する。 */
@Service
public class Mcm0011uService {
    @Autowired private JdbcTemplate jdbcTemplate;
    @Autowired private Mcm2004uService permissions;
    @Autowired private Validator validator;

    public static final String[] INPUT_FIELDS = {"nonyusakiCd", "nonyusakiNk", "kyunonyusakiNk",
        "nonyusakikojoNk", "nonyusakikanaKn", "nonyusakieimeiEn", "yubinNo", "jusyo1Nk", "jusyo2Nk",
        "kuniNk", "telNo", "faxNo", "biko"};
    private static final String TABLE = "MCM.MCM_MA_NONYUSAKI";
    private static final String CONFLICT = "他の利用者によって納入先が更新または削除されています。画面を開き直してください。";

    public boolean canView(String user, String screen) {
        return user != null && Set.of("1", "2").contains(Objects.toString(permissions.getAuthority(user, screen), "0"));
    }
    public boolean canUpdate(String user, String screen) {
        return user != null && "2".equals(permissions.getAuthority(user, screen));
    }
    private void requireUpdate(String user) {
        if (!canUpdate(user, "MCM0011U")) throw new IllegalStateException("更新権限がありません。");
    }
    @Transactional(readOnly = true)
    public Optional<NonyusakiEntity> findByNonyusakiId(BigDecimal id) {
        return jdbcTemplate.query("SELECT * FROM " + TABLE + " WHERE NONYUSAKI_ID = ?",
            BeanPropertyRowMapper.newInstance(NonyusakiEntity.class), id).stream().findFirst();
    }
    public Mcm0011uForm initNewForm() {
        Mcm0011uForm form = new Mcm0011uForm();
        form.setNonyusakiId(BigDecimal.ZERO);
        form.setDtsrenkeiFlg(BigDecimal.ZERO);
        form.setNewMode(true);
        return form;
    }
    public Mcm0011uForm entityToForm(NonyusakiEntity entity) {
        Mcm0011uForm form = new Mcm0011uForm();
        BeanUtils.copyProperties(entity, form);
        form.setRevision(revision(entity));
        form.setNewMode(false);
        return form;
    }
    private String revision(NonyusakiEntity entity) {
        try {
            return HexFormat.of().formatHex(MessageDigest.getInstance("SHA-256")
                .digest(entity.toString().getBytes(StandardCharsets.UTF_8)));
        } catch (java.security.NoSuchAlgorithmException ex) { throw new IllegalStateException(ex); }
    }
    private NonyusakiEntity current(Mcm0011uForm form) {
        var rows = jdbcTemplate.query("SELECT * FROM " + TABLE + " WITH (UPDLOCK, HOLDLOCK) WHERE NONYUSAKI_ID = ?",
            BeanPropertyRowMapper.newInstance(NonyusakiEntity.class), form.getNonyusakiId());
        if (rows.size() != 1 || !Objects.equals(form.getRevision(), revision(rows.get(0))))
            throw new IllegalStateException(CONFLICT);
        return rows.get(0);
    }
    private void validate(Mcm0011uForm form) {
        var errors = validator.validate(form);
        if (!errors.isEmpty()) throw new IllegalStateException(errors.stream().map(v -> v.getMessage())
            .distinct().sorted().collect(Collectors.joining("\n")));
    }
    public boolean hasChanges(Mcm0011uForm form, NonyusakiEntity entity) {
        var input = new BeanWrapperImpl(form);
        var original = new BeanWrapperImpl(entity);
        for (String field : INPUT_FIELDS) {
            if (!Objects.toString(input.getPropertyValue(field), "").equals(
                    Objects.toString(original.getPropertyValue(field), ""))) return true;
        }
        return false;
    }
    private String editor(String loginId) {
        var names = jdbcTemplate.queryForList("SELECT TANTO_NK FROM MCM.MCM_MO_TANTO WHERE LOGIN_ID = ?", String.class, loginId);
        return names.stream().filter(Objects::nonNull).filter(n -> !n.isBlank()).findFirst().orElse(loginId);
    }

    /** 全カラムの読込時値を比較し、採番・更新を同一トランザクションで行う。 */
    @Transactional(isolation = Isolation.SERIALIZABLE)
    public BigDecimal save(Mcm0011uForm form, String loginId) {
        requireUpdate(loginId);
        boolean insert = form.getNonyusakiId() == null || form.getNonyusakiId().signum() == 0;
        NonyusakiEntity original = insert ? null : current(form);
        if (!insert && !hasChanges(form, original)) return original.getNonyusakiId();
        validate(form);
        LocalDateTime now = LocalDateTime.now();
        String user = editor(loginId);
        if (insert) {
            BigDecimal id = jdbcTemplate.queryForObject("SELECT ISNULL(MAX(NONYUSAKI_ID), 0) + 1 FROM "
                + TABLE + " WITH (UPDLOCK, HOLDLOCK)", BigDecimal.class);
            jdbcTemplate.update("INSERT INTO " + TABLE + " (NONYUSAKI_ID, NONYUSAKI_CD, NONYUSAKI_NK, KYUNONYUSAKI_NK, "
                + "NONYUSAKIKOJO_NK, NONYUSAKIKANA_KN, NONYUSAKIEIMEI_EN, YUBIN_NO, JUSYO1_NK, JUSYO2_NK, "
                + "KUNI_NK, TEL_NO, FAX_NO, BIKO, DTSRENKEI_FLG, CREATED_DT, CREATED_BY, LASTUPDATE_DT, LASTUPDATE_BY) "
                + "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", id,
                form.getNonyusakiCd(), form.getNonyusakiNk(), form.getKyunonyusakiNk(), form.getNonyusakikojoNk(),
                form.getNonyusakikanaKn(), form.getNonyusakieimeiEn(), form.getYubinNo(), form.getJusyo1Nk(),
                form.getJusyo2Nk(), form.getKuniNk(), form.getTelNo(), form.getFaxNo(), form.getBiko(),
                BigDecimal.ZERO, now, user, now, user);
            return id;
        }
        int count = jdbcTemplate.update("UPDATE " + TABLE + " SET NONYUSAKI_CD = ?, NONYUSAKI_NK = ?, KYUNONYUSAKI_NK = ?, "
            + "NONYUSAKIKOJO_NK = ?, NONYUSAKIKANA_KN = ?, NONYUSAKIEIMEI_EN = ?, YUBIN_NO = ?, JUSYO1_NK = ?, JUSYO2_NK = ?, "
            + "KUNI_NK = ?, TEL_NO = ?, FAX_NO = ?, BIKO = ?, LASTUPDATE_DT = ?, LASTUPDATE_BY = ? WHERE NONYUSAKI_ID = ?",
            form.getNonyusakiCd(), form.getNonyusakiNk(), form.getKyunonyusakiNk(), form.getNonyusakikojoNk(),
            form.getNonyusakikanaKn(), form.getNonyusakieimeiEn(), form.getYubinNo(), form.getJusyo1Nk(),
            form.getJusyo2Nk(), form.getKuniNk(), form.getTelNo(), form.getFaxNo(), form.getBiko(), now, user, form.getNonyusakiId());
        if (count != 1) throw new IllegalStateException(CONFLICT);
        return form.getNonyusakiId();
    }

    /** VBと同じく削除時は入力必須検証を行わず、利用プラントと排他を確認する。 */
    @Transactional(isolation = Isolation.SERIALIZABLE)
    public boolean delete(Mcm0011uForm form, String loginId) {
        requireUpdate(loginId);
        if (form.getNonyusakiId() == null || form.getNonyusakiId().signum() <= 0)
            throw new IllegalStateException("未登録の納入先は削除できません。");
        current(form);
        if (countPlantByNonyusakiId(form.getNonyusakiId()) > 0) return false;
        if (jdbcTemplate.update("DELETE FROM " + TABLE + " WHERE NONYUSAKI_ID = ?", form.getNonyusakiId()) != 1)
            throw new IllegalStateException(CONFLICT);
        return true;
    }
    public int countPlantByNonyusakiId(BigDecimal id) {
        Integer count = jdbcTemplate.queryForObject("SELECT COUNT(PLANT_ID) FROM MCM.MCM_MA_PLANT WHERE NONYUSAKI_ID = ?", Integer.class, id);
        return count == null ? 0 : count;
    }
}
