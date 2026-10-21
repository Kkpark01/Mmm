package com.daifuku.mcm.dto;

/**
 * 【変換元】CPMenuData (FRAMEWORK/COMPONENT/CPMenuData.vb)
 * 
 * メニュー項目DTO
 * VB.NETのCPMenuDataクラスを置き換えるDTO。
 * メニューID、メニュー名、メニュー区分（menuDivision）、利用権限を保持する。
 */
public class MenuItemDto {

    /** 画面ID（例: MCM0010U） */
    private String menuId;

    /** 画面表示名（例: 納入先・プラント検索） */
    private String menuName;

    /**
     * メニュー区分（menuDivision）
     * 【変換元】McmMenuButton.vb - MenuDivision プロパティ
     *   元コード: Private _menuDivision As String
     * 
     * VB.NETではmenuDivision単位でフォームインスタンスを管理していた。
     * Web版ではURL routingに置き換えるため参考情報として保持。
     * 
     *   "10" = マスタ系
     *   "20" = 需要家系（取引先）
     *   "30" = カスタマー系（店舗）
     *   "40" = その他
     */
    private String menuKbn;

    /**
     * 利用権限
     * 【変換元】CPMenuData.Riyokengen
     *   0 = 全権限, 1 = 参照のみ（AUTHORITY_DIVISION_INSPECTION）
     */
    private int riyokengen;

    public MenuItemDto() {
    }

    public MenuItemDto(String menuId, String menuName, String menuKbn, int riyokengen) {
        this.menuId = menuId;
        this.menuName = menuName;
        this.menuKbn = menuKbn;
        this.riyokengen = riyokengen;
    }

    /**
     * 画面遷移先URLを返す。
     * 【変換元】McmMenuButton.vb - OnClick() → McmMenuForm.CallTargetScreen()
     *   元コード: menuForm.CallTargetScreen(Me.MenuDivision, Me.FunctionID)
     *   Web版ではmenuIdを小文字化したパスに遷移する。
     *
     * @return 遷移先URL（例: /mcm0010u）
     */
    public String getUrl() {
        // メニューで開く時だけ初期化する。詳細画面から通常URLへ戻る場合は検索を保持。
        if ("MCM1003U".equalsIgnoreCase(menuId)) return "/mcm1003u/open";
        if ("MCM3004U".equalsIgnoreCase(menuId)) return "/mcm3004u/open";
        return "/" + menuId.toLowerCase();
    }

    // --- Getters and Setters ---

    public String getMenuId() {
        return menuId;
    }

    public void setMenuId(String menuId) {
        this.menuId = menuId;
    }

    public String getMenuName() {
        return menuName;
    }

    public void setMenuName(String menuName) {
        this.menuName = menuName;
    }

    public String getMenuKbn() {
        return menuKbn;
    }

    public void setMenuKbn(String menuKbn) {
        this.menuKbn = menuKbn;
    }

    public int getRiyokengen() {
        return riyokengen;
    }

    public void setRiyokengen(int riyokengen) {
        this.riyokengen = riyokengen;
    }
}
