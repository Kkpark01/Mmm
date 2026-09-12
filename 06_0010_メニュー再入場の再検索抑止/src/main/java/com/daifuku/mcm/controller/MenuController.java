package com.daifuku.mcm.controller;

import java.util.List;

import jakarta.servlet.http.HttpSession;

import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;

/**
 * 【変換元】McmMenuForm.vb (約480行) + Mcm0001uScreen.vb (約350行)
 *
 * メニュー画面コントローラー（更新版）
 * SuccessHandler によりセッションにUserInfoがセットされている前提。
 */

import com.daifuku.mcm.common.AppConstants;
import com.daifuku.mcm.common.BaseController;
import com.daifuku.mcm.common.CustomerScreenSupport;
import com.daifuku.mcm.dto.MenuGroup;
import com.daifuku.mcm.dto.UserInfo;
import com.daifuku.mcm.service.MenuService;

@Controller
@RequestMapping("/menu")
public class MenuController extends BaseController {

    private final MenuService menuService;

    public MenuController(MenuService menuService) {
        this.menuService = menuService;
    }

    @Override
    protected String getScreenTitle() {
        return "メニュー";
    }

    @Override
    protected String getFunctionId() {
        return "MCM0001U";
    }

    /**
     * 【変換元】Mcm0001uScreen.vb - Mcm0001uScreen_Load()
     */
    @GetMapping
    public String index(Model model, HttpSession session) {

        // ★ MCM関連のセッションをクリア
        clearMcmSession(session);

        // 検索→詳細→戻るでは0010の条件を保持するが、メニューへ戻ったら検索を終了する。
        // 結果・削除用の比較情報も解放し、再入場時の自動再検索と古い画面からの操作を防ぐ。
        synchronized (session) {
            for (String key : List.of("mcm0010u_search_form", "mcm0010u_owner", "mcm0010u_nonyusaki",
                    "mcm0010u_plants", "mcm0010u_brands", "mcm0010u_kiki", "mcm0010u_versions")) {
                session.removeAttribute(key);
            }
            CustomerScreenSupport.rotate(session, "MCM0010U");
        }
    	
        // セッションからユーザー情報を取得（SuccessHandlerでセット済み）
        UserInfo userInfo = getLoginUserInfo(session);
        String authorityDivision = (String) session.getAttribute(
                AppConstants.SESSION_AUTHORITY_DIVISION);

        if (userInfo == null) {
            return "redirect:/login";
        }

        // 【追加 #144】システム管理関連カテゴリの権限をセッションから取得
        String systemManagementKengen = (String) session.getAttribute(
                AppConstants.SESSION_SYSTEM_MANAGEMENT_KENGEN);

        // メニューデータ取得（権限フィルタリング済み）
        List<MenuGroup> menuGroups = menuService.getMenuGroups(
                userInfo.getUserId(), authorityDivision, systemManagementKengen);

        model.addAttribute("menuGroups", menuGroups);
        model.addAttribute("userName", userInfo.getUserName());
        model.addAttribute("userId", userInfo.getUserId());
        model.addAttribute("jigyosyoName", userInfo.getJigyosyoName());

        // 共通属性セット
        setCommonAttributes(model, session);

        return "menu/index";
    }

    /**
     * 【変換元】McmMenuForm.vb - McmMenuForm_FormClosing()
     */
    @GetMapping("/logout")
    public String logout(HttpSession session) {
        session.invalidate();
        return "redirect:/login";
    }
}
