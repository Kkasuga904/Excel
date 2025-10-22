import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import ChatMessage from './components/ChatMessage';
import ChatInput from './components/ChatInput';
import LoadingSpinner from './components/LoadingSpinner';
import './taskpane.css';

declare const Office: any;
declare const Excel: any;

interface Message {
  id: string;
  text: string;
  sender: 'user' | 'ai';
  timestamp: Date;
  isError?: boolean;
  isSuccess?: boolean;
}

interface RangeData {
  values: unknown[][];
  address: string;
  formulas: unknown[][];
}

interface HealthStatus {
  state: 'loading' | 'success' | 'error';
  message: string;
}

interface TaskPaneProps {
  healthStatus?: HealthStatus;
}

const TaskPane: React.FC<TaskPaneProps> = ({ healthStatus }) => {
  const [messages, setMessages] = useState<Message[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [officeInitialized, setOfficeInitialized] = useState(false);
  const [isStandaloneMode, setIsStandaloneMode] = useState(false);
  const [hostName, setHostName] = useState<string | null>(null);
  const standaloneNoticeShown = useRef(false);
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const apiBase = useMemo(() => process.env.REACT_APP_API_BASE_URL ?? '', []);

  const buildUrl = (path: string) => {
    const normalizedBase = apiBase.replace(/\/+$/, '');
    return `${normalizedBase}${path}`;
  };

  const scrollToBottom = () => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  };

  useEffect(() => {
    scrollToBottom();
  }, [messages]);

  const addMessage = useCallback(
    (
      text: string,
      sender: 'user' | 'ai',
      isError: boolean = false,
      isSuccess: boolean = false
    ) => {
      setMessages((prev) => [
        ...prev,
        {
          id: Date.now().toString(),
          text,
          sender,
          timestamp: new Date(),
          isError,
          isSuccess
        }
      ]);
    },
    []
  );

  useEffect(() => {
    let cancelled = false;

    const initOffice = async () => {
      if (typeof Office === 'undefined' || typeof Office.onReady !== 'function') {
        console.warn('Office.js が見つかりません。Excel 以外の環境で開かれています。');
        if (!cancelled) {
          setIsStandaloneMode(true);
          setHostName(null);
          setOfficeInitialized(true);
        }
        return;
      }

      try {
        const info = await Office.onReady();
        if (cancelled) {
          return;
        }

        const detectedHost = typeof info?.host === 'string' ? info.host : null;
        setHostName(detectedHost);

        if (detectedHost && detectedHost.toLowerCase() === 'excel') {
          setIsStandaloneMode(false);
        } else {
          console.warn(
            `Office.js は Excel 以外のホスト (${detectedHost ?? '不明'}) で読み込まれました。スタンドアロンモードで起動します。`
          );
          setIsStandaloneMode(true);
        }
      } catch (error) {
        if (!cancelled) {
          console.error('Office.js の初期化に失敗しました。スタンドアロンモードで継続します。', error);
          setIsStandaloneMode(true);
          setHostName(null);
          addMessage(
            'Office.js の初期化に失敗しました。ブラウザ単体では Excel の機能は利用できませんが、チャットは継続できます。',
            'ai',
            true
          );
        }
      } finally {
        if (!cancelled) {
          setOfficeInitialized(true);
        }
      }
    };

    void initOffice();

    return () => {
      cancelled = true;
    };
  }, [addMessage]);

  const getSelectedData = async (): Promise<RangeData> => {
    if (typeof Excel === 'undefined' || typeof Excel.run !== 'function') {
      throw new Error('Excel 対応の環境ではありません。');
    }

    try {
      return await Excel.run(async (context: any) => {
        const range = context.workbook.getSelectedRange();
        range.load('values, address, formulas');
        await context.sync();
        return {
          values: range.values,
          address: range.address,
          formulas: range.formulas
        } as RangeData;
      });
    } catch (error) {
      console.error('Failed to get selected data:', error);
      throw new Error('選択範囲の取得に失敗しました。');
    }
  };

  const handleSendMessage = async (userMessage: string) => {
    addMessage(userMessage, 'user');
    setIsLoading(true);

    try {
      let cellData: RangeData | null = null;
      let abortRequest = false;

      if (!isStandaloneMode) {
        try {
          cellData = await getSelectedData();
        } catch (error) {
          console.warn('Selection read failed:', error);
          addMessage(
            'セル範囲を取得できませんでした。セルを選択してからもう一度お試しください。',
            'ai',
            true
          );
          abortRequest = true;
        }
      } else if (!standaloneNoticeShown.current) {
        addMessage(
          'Excel 以外の環境ではセルの内容を取得できませんが、チャットは利用できます。',
          'ai'
        );
        standaloneNoticeShown.current = true;
      }

      if (abortRequest) {
        return;
      }

      const response = await fetch(buildUrl('/api/chat'), {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          message: userMessage,
          cellData,
          messageHistory: messages.map((m) => ({
            role: m.sender === 'user' ? 'user' : 'assistant',
            content: m.text
          }))
        })
      });

      if (!response.ok) {
        const errorData = await response.json().catch(() => undefined);
        throw new Error(errorData?.error || 'API へのリクエストに失敗しました。');
      }

      const result = await response.json();
      addMessage(result.message, 'ai', false, result.action !== 'none');

      if (result.action === 'write' && result.data) {
        if (isStandaloneMode) {
          addMessage('現在の環境ではセルへの書き込みは行えません。Excel 上で実行してください。', 'ai', true);
        } else {
          try {
            await Excel.run(async (context: any) => {
              const sheet = context.workbook.worksheets.getActiveWorksheet();
              const range = sheet.getRange(result.data.address);
              range.values = result.data.values;
              await context.sync();
            });

            addMessage(`${result.data.address} に結果を書き込みました。`, 'ai', false, true);
          } catch (error) {
            console.error('Failed to write to cell:', error);
            addMessage('セルへの書き込みに失敗しました。Excel 上で再度お試しください。', 'ai', true);
          }
        }
      }
    } catch (error) {
      console.error('Error:', error);
      const message = error instanceof Error ? error.message : 'エラーが発生しました。';
      addMessage(message, 'ai', true);
    } finally {
      setIsLoading(false);
    }
  };

  if (!officeInitialized) {
    return (
      <div className="chat-container">
        <div className="chat-messages">
          <LoadingSpinner message="Office.js を初期化しています..." />
        </div>
      </div>
    );
  }

  return (
    <div className="taskpane-root">
      <div className="chat-shell">
        <header className="chat-header">
          <h1>Excel AI チャットアシスタント</h1>
          <p>Excel の選択範囲と会話しながら作業を進めるための補助ツールです。</p>
        </header>
        {healthStatus && (
          <div className={`status-banner status-${healthStatus.state}`}>
            {healthStatus.message}
          </div>
        )}
        {isStandaloneMode && (
          <div className="standalone-banner">
            Excel 以外の環境{hostName ? `（${hostName}）` : ''}で動作しています。セルの読み取り・書き込みは無効ですが、チャットは利用できます。
          </div>
        )}
        <div className="chat-container">
          <div className="chat-messages">
            {messages.length === 0 ? (
              <div className="empty-state">
                <div className="empty-state-icon">💬</div>
                <div className="empty-state-title">Excel AI チャットアシスタント</div>
                <div className="empty-state-description">
                  Excel でセル範囲を選択してから質問すると、選択中のデータを基に回答します。
                  <br />
                  ブラウザ単体ではセルの取得はできませんが、チャットでの相談が可能です。
                </div>
              </div>
            ) : (
              <>
                {messages.map((msg) => (
                  <ChatMessage
                    key={msg.id}
                    message={msg.text}
                    sender={msg.sender}
                    timestamp={msg.timestamp}
                    isError={msg.isError}
                    isSuccess={msg.isSuccess}
                  />
                ))}
                {isLoading && <LoadingSpinner message="考えています..." />}
                <div ref={messagesEndRef} />
              </>
            )}
          </div>
          <ChatInput
            onSendMessage={handleSendMessage}
            isLoading={isLoading}
            placeholder="例: この表を要約して"
          />
        </div>
      </div>
      <details className="info-panel">
        <summary>使い方とヒント</summary>
        <ul>
          <li>Excel でセル範囲を選択すると、そのデータをコンテキストに回答します。</li>
          <li>「表を整形して」「グラフを作成して」などの指示で具体的な操作案を得られます。</li>
          <li>ブラウザでのテストを終えたら README の手順で sideload し、本番環境で確認してください。</li>
        </ul>
      </details>
    </div>
  );
};

export default TaskPane;
