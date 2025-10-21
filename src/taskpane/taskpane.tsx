import React, { useState, useEffect, useRef } from 'react';
import ReactDOM from 'react-dom/client';
import ChatMessage from './components/ChatMessage';
import ChatInput from './components/ChatInput';
import LoadingSpinner from './components/LoadingSpinner';
import './taskpane.css';

interface Message {
  id: string;
  text: string;
  sender: 'user' | 'ai';
  timestamp: Date;
  isError?: boolean;
  isSuccess?: boolean;
}

const TaskPane: React.FC = () => {
  const [messages, setMessages] = useState<Message[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [selectedData, setSelectedData] = useState<any>(null);
  const messagesEndRef = useRef<HTMLDivElement>(null);

  // メッセージを自動スクロール
  const scrollToBottom = () => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  };

  useEffect(() => {
    scrollToBottom();
  }, [messages]);

  // Office.jsの初期化
  useEffect(() => {
    const initOffice = async () => {
      try {
        await Office.onReady();
        console.log('Office.js initialized');
      } catch (error) {
        console.error('Office.js initialization failed:', error);
        addMessage(
          'Office.jsの初期化に失敗しました。ブラウザを再読み込みしてください。',
          'ai',
          true
        );
      }
    };

    initOffice();
  }, []);

  // メッセージを追加
  const addMessage = (
    text: string,
    sender: 'user' | 'ai',
    isError: boolean = false,
    isSuccess: boolean = false
  ) => {
    const newMessage: Message = {
      id: Date.now().toString(),
      text,
      sender,
      timestamp: new Date(),
      isError,
      isSuccess
    };
    setMessages((prev) => [...prev, newMessage]);
  };

  // 選択されたセルデータを取得
  const getSelectedData = async () => {
    try {
      return await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load('values, address, formulas');
        await context.sync();
        return {
          values: range.values,
          address: range.address,
          formulas: range.formulas
        };
      });
    } catch (error) {
      console.error('Failed to get selected data:', error);
      throw new Error('セルデータの取得に失敗しました');
    }
  };

  // メッセージ送信
  const handleSendMessage = async (userMessage: string) => {
    // ユーザーメッセージを追加
    addMessage(userMessage, 'user');
    setIsLoading(true);

    try {
      // 選択されたセルデータを取得
      let cellData = null;
      try {
        cellData = await getSelectedData();
        setSelectedData(cellData);
      } catch (error) {
        console.warn('Could not get selected data:', error);
        addMessage(
          'データが選択されていません。セルを選択してからもう一度お試しください。',
          'ai',
          true
        );
        setIsLoading(false);
        return;
      }

      // バックエンドにリクエスト送信
      const response = await fetch('http://localhost:3001/api/chat', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          message: userMessage,
          cellData: cellData,
          messageHistory: messages.map((m) => ({
            role: m.sender === 'user' ? 'user' : 'assistant',
            content: m.text
          }))
        })
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error || 'API呼び出しに失敗しました');
      }

      const result = await response.json();

      // AIの返答を表示
      addMessage(result.message, 'ai', false, result.action !== 'none');

      // セルに結果を書き込み
      if (result.action === 'write' && result.data) {
        try {
          await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRange(result.data.address);
            range.values = result.data.values;
            await context.sync();
          });

          addMessage(
            `${result.data.address}に結果を書き込みました`,
            'ai',
            false,
            true
          );
        } catch (error) {
          console.error('Failed to write to cell:', error);
          addMessage(
            'セルへの書き込みに失敗しました。手動で入力してください。',
            'ai',
            true
          );
        }
      }
    } catch (error) {
      console.error('Error:', error);
      const errorMessage =
        error instanceof Error ? error.message : 'エラーが発生しました';
      addMessage(errorMessage, 'ai', true);
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <div className="chat-container">
      <div className="chat-messages">
        {messages.length === 0 ? (
          <div className="empty-state">
            <div className="empty-state-icon">💬</div>
            <div className="empty-state-title">Excel AI チャットアシスタント</div>
            <div className="empty-state-description">
              Excelのセルを選択して、自然言語で指示してください。
              <br />
              データ分析、操作、レポート作成などが可能です。
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
            {isLoading && <LoadingSpinner message="処理中..." />}
            <div ref={messagesEndRef} />
          </>
        )}
      </div>
      <ChatInput
        onSendMessage={handleSendMessage}
        isLoading={isLoading}
        placeholder="例: このデータを分析して"
      />
    </div>
  );
};

// React アプリケーションをマウント
const root = ReactDOM.createRoot(document.getElementById('root')!);
root.render(<TaskPane />);

