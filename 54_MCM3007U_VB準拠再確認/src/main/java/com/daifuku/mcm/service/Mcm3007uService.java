package com.daifuku.mcm.service;

import com.daifuku.mcm.exception.McmBusinessException;
import com.daifuku.mcm.form.Mcm3007uForm;
import com.daifuku.mcm.repository.Mcm3007uRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Isolation;
import org.springframework.transaction.annotation.Transactional;
import java.math.BigDecimal;
import java.time.YearMonth;
import java.util.*;

@Service
public class Mcm3007uService {
    @Autowired private Mcm3007uRepository repository;

    public void search(Mcm3007uForm f) {
        var um = new ArrayList<String>();
        var tk = new ArrayList<String>();
        if (f.isJotaiUmMitsumori()) um.add("1");
        if (f.isJotaiUmKeiyaku()) um.add("2");
        if (f.isJotaiUmKaiyaku()) um.add("4");
        if (f.isJotaiUmHaki()) um.add("3");
        if (f.isJotaiTkMitsumori()) tk.add("1");
        if (f.isJotaiTkKeiyaku()) tk.add("2");
        if (f.isJotaiTkKaiyaku()) tk.add("4");
        if (f.isJotaiTkHaki()) tk.add("3");
        if (um.isEmpty() && tk.isEmpty() && java.util.stream.Stream.of(
                f.getNonyusakiCd(), f.getNonyusakiNk(), f.getSupportId(), f.getPlantNk())
                .allMatch(s -> s == null || s.isBlank())) {
            throw new McmBusinessException("検索条件を指定してください。");
        }
        // 途中の検索失敗で、一部の表だけ新しい結果になることを防ぐ。
        var plants = repository.searchNonyusaki(f.getNonyusakiCd(), f.getNonyusakiNk(),
                f.getSupportId(), f.getPlantNk(), um);
        if (plants.isEmpty()) {
            f.setGrid1Rows(List.of()); f.setGrid2Rows(List.of()); f.setGrid3Rows(List.of());
            return;
        }
        var quotes = repository.searchUva(f.getNonyusakiCd(), f.getNonyusakiNk(),
                f.getSupportId(), f.getPlantNk(), um);
        var suppliers = repository.searchTka(f.getNonyusakiCd(), f.getNonyusakiNk(),
                f.getSupportId(), f.getPlantNk(), tk);
        f.setGrid1Rows(plants); f.setGrid2Rows(quotes); f.setGrid3Rows(suppliers);
    }

    public boolean hasBrand(BigDecimal plantId) { return repository.hasBrand(plantId); }

    @Transactional(isolation = Isolation.SERIALIZABLE)
    public void discard(boolean um, Map<String,String> row, String userName) {
        BigDecimal contract = Mcm3007uRepository.id(row, um ? "UK_KEIYAKU_ID" : "TK_KEIYAKU_ID");
        BigDecimal source = Mcm3007uRepository.id(row, um ? "UM_KIHON_MITSUMORI_ID" : "TM_KEIYAKUJIKAN_ID");
        if (contract == null || source == null || contract.signum() <= 0 || source.signum() <= 0
                || row.getOrDefault("KEIYAKU_DEL", "").isBlank()) {
            throw new McmBusinessException("破棄できない行です。再検索してください。");
        }
        if (userName == null || userName.isBlank()) {
            throw new McmBusinessException("ログイン利用者情報を確認できません。再ログインしてください。");
        }
        // 再検索後の変更も、更新トランザクション内で契約・元見積をロックして検出する。
        repository.verifyDiscard(um, contract, source, row);
        if (!um) {
            var min = repository.getMinSiharaiTsuki(contract);
            if (min != null && YearMonth.from(min).isBefore(repository.currentMonth())) {
                throw new McmBusinessException("既に支払い明細が存在するため、この取引先契約を破棄できません。");
            }
        }
        repository.discard(um, contract, source, row, userName);
    }
}
