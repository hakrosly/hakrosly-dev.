Attribute VB_Name = "modConst"
Option Explicit
' ============================================================
' modConst : 名前や列の番号など、ツール全体で使う決まった値
' ============================================================

Public Const TOOL_VERSION As String = "0.1.0"
Public Const ERR_TOOL As Long = 20001          ' このツールが出すエラーの番号

' ---- ツール用ブックのシートの名前 ----
Public Const SH_MAIN As String = "操作"
Public Const SH_CFG_BOM As String = "設定_部品表"
Public Const SH_CFG_MTX As String = "設定_マトリックス"
Public Const SH_CFG_TXT As String = "設定_文字"
Public Const SH_MAP As String = "仕様対応表"
Public Const SH_RD_BOM As String = "読取_部品表"
Public Const SH_RD_MTX As String = "読取_マトリックス"
Public Const SH_RD_WARN As String = "読取_警告"
Public Const SH_STATE As String = "読取_記録"
Public Const SH_TEST As String = "テスト結果"

' ---- 結果ブックのシートの名前 ----
Public Const R_SUM As String = "まとめ"
Public Const R_LIST As String = "食い違い一覧"
Public Const R_MTX As String = "チェック済み_マトリックス"
Public Const R_CELLS As String = "チェック済み_セル一覧"
Public Const R_BOM As String = "チェック済み_部品表"
Public Const R_PROOF As String = "検算"
Public Const R_SAMPLE As String = "抜き取り確認"
Public Const R_CFG As String = "設定の控え"
Public Const RESULT_FOLDER As String = "照合結果"

' ---- 値を表す言葉 ----
Public Const V_YES As String = "あり"
Public Const V_NO As String = "なし"
Public Const V_UNK As String = "判断できない"

' ---- 状態を表す言葉 (消し込みに使う) ----
Public Const S_OK As String = "一致"
Public Const S_NG As String = "不一致"
Public Const S_UNK As String = "判断できない"
Public Const S_OUT As String = "対象外"

' ---- 出どころ ----
Public Const SRC_BOM As String = "部品表"
Public Const SRC_MTX As String = "マトリックス"

' ---- 設定に書く言葉 ----
Public Const O_VERT As String = "縦"
Public Const O_HORZ As String = "横"
Public Const K_MARK As String = "印"
Public Const K_QTY As String = "員数"
Public Const OPT_YES As String = "する"
Public Const OPT_NO As String = "しない"
Public Const OPT_SPLIT As String = "分ける"
Public Const OPT_NOSPLIT As String = "分けない"
Public Const MAP_EXCLUDE As String = "対象外"

' ---- 設定_文字 の「種類」 ----
Public Const T_HYPHEN As String = "ハイフンとみなす文字"
Public Const T_YES As String = "あり を表す印"
Public Const T_NO As String = "なし を表す印"
Public Const T_IGNORE As String = "ハイフンを無視する"

' ---- 読取_部品表 の列 ----
Public Const BC_KEY As Long = 1        ' 品番キー (比べるための形)
Public Const BC_PART As Long = 2       ' 品番 (元の文字)
Public Const BC_NAME As Long = 3       ' 名称
Public Const BC_SPECKEY As Long = 4    ' 仕様コード (比べるための形)
Public Const BC_SPEC As Long = 5       ' 仕様コード (元の文字)
Public Const BC_QTY As Long = 6        ' 員数
Public Const BC_INSCOPE As Long = 7    ' 範囲の中 (○ か空)
Public Const BC_ROW As Long = 8        ' 元の行
Public Const BC_N As Long = 8

' ---- 読取_マトリックス の列 ----
Public Const MC_KEY As Long = 1        ' 品番キー
Public Const MC_PART As Long = 2       ' 品番 (元の文字)
Public Const MC_NAME As Long = 3       ' 名称
Public Const MC_HEAD As Long = 4       ' 見出しの文字
Public Const MC_VAL As Long = 5        ' 値 (あり / なし / 判断できない)
Public Const MC_QTY As Long = 6        ' 員数
Public Const MC_RAW As Long = 7        ' セルに入っていた文字
Public Const MC_PARTCELL As Long = 8   ' 品番のセル
Public Const MC_VALCELL As Long = 9    ' 値のセル
Public Const MC_NAMECELL As Long = 10  ' 名称のセル
Public Const MC_HEADCELL As Long = 11  ' 見出しのセル
Public Const MC_N As Long = 11

' ---- 読取_警告 の列 ----
Public Const WC_SRC As Long = 1        ' 出どころ (部品表 / マトリックス)
Public Const WC_CODE As Long = 2       ' 記号 (W-1 など)
Public Const WC_PART As Long = 3       ' 品番
Public Const WC_HEAD As Long = 4       ' 見出しの文字
Public Const WC_CELL As Long = 5       ' セル
Public Const WC_RAW As Long = 6        ' 元の文字
Public Const WC_NOTE As Long = 7       ' 説明
Public Const WC_N As Long = 7

' ---- 仕様対応表 の列 ----
Public Const PC_LAYOUT As Long = 1     ' 設定名
Public Const PC_HEAD As Long = 2       ' 見出しの文字
Public Const PC_SPEC As Long = 3       ' 仕様コード
Public Const PC_EXCL As Long = 4       ' 対象外
Public Const PC_MEMO As Long = 5       ' メモ
Public Const PC_N As Long = 5

' ---- 食い違い一覧 の列 ----
Public Const FC_CLASS As Long = 1      ' 区分 (不一致 / 警告)
Public Const FC_CODE As Long = 2       ' 記号
Public Const FC_TITLE As Long = 3      ' 名前
Public Const FC_PART As Long = 4       ' 品番
Public Const FC_NAME_B As Long = 5     ' 名称 (部品表)
Public Const FC_NAME_M As Long = 6     ' 名称 (マトリックス)
Public Const FC_SPEC As Long = 7       ' 仕様コード
Public Const FC_HEAD As Long = 8       ' 見出しの文字
Public Const FC_VAL_B As Long = 9      ' 部品表の値
Public Const FC_VAL_M As Long = 10     ' マトリックスの値
Public Const FC_CELL As Long = 11      ' マトリックスのセル
Public Const FC_NOTE As Long = 12      ' 説明
Public Const FC_N As Long = 12

Public Const CLASS_NG As String = "不一致"
Public Const CLASS_WARN As String = "警告"

' ---- 色 (結果ブックで使う) ----
Public Const COL_OK As Long = 13561798      ' 緑 RGB(198,239,206)
Public Const COL_NG As Long = 13551615      ' 赤 RGB(255,199,206)
Public Const COL_UNK As Long = 10284031     ' 黄 RGB(255,235,156)
Public Const COL_OUT As Long = 14277081     ' 灰 RGB(217,217,217)
Public Const COL_HEAD As Long = 16247773    ' 水色 RGB(221,235,247)

' ---- 検算のあと、わざと値を変えて確かめる回数 ----
Public Const MUTATION_COUNT As Long = 10
' ---- 抜き取り確認に出す件数 ----
Public Const SAMPLE_COUNT As Long = 10
