```mermaid
flowchart TD
    classDef process fill:#d4f1f9,stroke:#333,stroke-width:1px;
    classDef decision fill:#ffe6cc,stroke:#333,stroke-width:1px;
    classDef start fill:#d5e8d4,stroke:#333,stroke-width:1px;
    classDef endNode fill:#f8cecc,stroke:#333,stroke-width:1px;

    %% メインフロー - メニュー作成
    A[onOpen] -->|スプレッドシート起動時| B[メニュー作成：🔍検索ツール]
    B --> C[ユーザーがメニュー選択]
    
    %% 設定ダイアログ
    C --> D[showSettingsDialog]
    D --> E{サブフォルダを\n検索しますか？}
    E -->|はい| F[executeSearchWithOption\nwith searchSubfolders=true]
    E -->|いいえ| G[executeSearchWithOption\nwith searchSubfolders=false]
    
    %% 検索フロー
    subgraph search [検索プロセス]
        H[キーワードとフォルダID取得] --> I{キーワードが\n存在する？}
        I -->|いいえ| J[エラーメッセージ表示]:::endNode
        I -->|はい| K[結果欄をクリア]
        
        K --> L{フォルダIDが\n存在する？}
        L -->|いいえ| M[エラーメッセージ表示]:::endNode
        L -->|はい| N[ファイル数・シート数カウント]
        
        N --> O[検索処理開始]
        O --> P[getAllFilesWithParent\nでファイル取得]
        P --> Q[各ファイルのシートを検索]
        Q --> R[各セルをキーワードで検索]
        R --> S[一致したセルを結果配列に追加]
        S --> T[進捗状況を更新]
        T --> U{すべて検索\n完了？}
        U -->|いいえ| O
        U -->|はい| V{結果が\n存在する？}
        
        V -->|はい| W[結果をスプレッドシートに表示]:::endNode
        V -->|いいえ| X[該当データなしメッセージ表示]:::endNode
    end
    
    %% ファイル取得処理
    subgraph files [ファイル取得処理]
        Y[getAllFilesWithParent] --> Z[指定フォルダのスプレッドシート\nファイルを取得]
        Z --> AA{サブフォルダも\n検索する？}
        AA -->|はい| AB[サブフォルダを再帰的に検索]
        AA -->|いいえ| AC[結果を返す]
        AB --> AC
    end
    
    F --> H
    G --> H
    P -.-> Y
    
    %% スタイル適用
    class A,B,C,D,H,K,N,O,P,Q,R,S,T,Y,Z,AB,AC process
    class E,I,L,U,V,AA decision
    class W,X,J,M endNode
```

## プロセスの説明

### 1. 初期設定

- **onOpen**: スプレッドシート開始時に実行され、カスタムメニュー「🔍検索ツール」を作成します
- **showSettingsDialog**: ユーザーがサブフォルダも検索するか選択するダイアログを表示します

### 2. 検索プロセス

- **executeSearchWithOption**: メイン検索処理を実行します
  - キーワードとフォルダIDの検証
  - 検索対象のファイル数とシート数を計算
  - 各ファイルとシートに対して検索実行
  - 進捗状況の定期的な更新
  - 結果の表示

### 3. ファイル収集

- **getAllFilesWithParent**: 指定されたフォルダから検索対象のスプレッドシートを収集
  - 再帰的にサブフォルダも検索（オプション選択時）
  - 親フォルダ情報も保持

検索は複数のフォルダを横断して行われ、指定したキーワードを含むセルをすべて検出して結果を一覧表示します。