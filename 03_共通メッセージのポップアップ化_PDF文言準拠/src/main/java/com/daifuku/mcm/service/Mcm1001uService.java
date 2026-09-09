/**
 * 【変換元】Mcm1001uScreen.vb - SearchButton_Click() / SearchNonyusaki() / SearchPlant()
 *   元ファイル行数: 約4,296行（Screen.vb全体）
 *   MCM1001U（取引先見積依頼作成検索）サービスクラス
 *
 * @author MCM Migration Tool
 * @since v8
 */
package com.daifuku.mcm.service;

import com.daifuku.mcm.dto.Mcm1001uNonyusakiDto;
import com.daifuku.mcm.dto.Mcm1001uPlantDto;
import com.daifuku.mcm.repository.Mcm1001uRepository;
import org.springframework.stereotype.Service;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

/**
 * MCM1001U 取引先見積依頼作成検索 サービス.
 *
 * 【変換元】Mcm1001uScreen.vb
 *   - SearchButton_Click() → search()
 *   - DataGridView_NONYUSAKI_SelectionChanged → getPlantByNonyusakiId()
 */
@Service
public class Mcm1001uService {

    private final Mcm1001uRepository repository;

    public Mcm1001uService(Mcm1001uRepository repository) {
        this.repository = repository;
    }

    // ========================================================
    // バリデーション
    // 【変換元】Mcm1001uScreen.vb - SearchButton_Click() 内の入力チェック
    // ========================================================

    public List<String> validate(String nonyusakiCd, String supportId,
            String nonyusakiNk, String plantNk, String etc) {
        List<String> errors = new ArrayList<>();
        if (isEmpty(nonyusakiCd) && isEmpty(supportId)
                && isEmpty(nonyusakiNk) && isEmpty(plantNk) && isEmpty(etc)) {
            errors.add("検索条件は、1項目以上選択して下さい。");
        }
        return errors;
    }

    // ========================================================
    // 納入先検索
    // 【変換元】Mcm1001uScreen.vb - SearchNonyusaki()
    //   ※ LIKEパラメータに%ワイルドカードを付与してRepositoryに渡す
    // ========================================================

    public List<Mcm1001uNonyusakiDto> searchNonyusaki(
            String nonyusakiCd, String supportId,
            String nonyusakiNk, String plantNk, String etc) {

        String cdParam = addWildcard(nonyusakiCd);
        String sidParam = addWildcard(supportId);
        String pnkParam = addWildcard(plantNk);

        return repository.searchNonyusaki(cdParam, sidParam, nonyusakiNk, pnkParam, etc);
    }

    // ========================================================
    // プラント検索（納入先ID指定）
    // 【変換元】Mcm1001uScreen.vb - DataGridView_NONYUSAKI_SelectionChanged
    // ========================================================

    public List<Mcm1001uPlantDto> getPlantByNonyusakiId(BigDecimal nonyusakiId) {
        return repository.searchPlant(nonyusakiId);
    }

    // ========================================================
    // プラント検索（検索条件付き）
    // ========================================================

    public List<Mcm1001uPlantDto> searchPlant(
            String nonyusakiCd, String supportId,
            String nonyusakiNk, String plantNk, String etc) {

        String cdParam = addWildcard(nonyusakiCd);
        String sidParam = addWildcard(supportId);
        String pnkParam = addWildcard(plantNk);

        return repository.searchPlantWithConditions(cdParam, sidParam, nonyusakiNk, pnkParam, etc);
    }

    // ========================================================
    // ユーティリティ
    // ========================================================

    private String addWildcard(String value) {
        if (isEmpty(value)) {
            return null;
        }
        return "%" + value.trim() + "%";
    }

    private static boolean isEmpty(String str) {
        return str == null || str.trim().isEmpty();
    }
}
