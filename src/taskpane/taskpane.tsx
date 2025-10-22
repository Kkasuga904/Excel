declare const Office: any;
declare const Excel: any;
import React, { useEffect, useMemo, useRef, useState } from 'react';
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

type RangeData = Excel.Interfaces.RangeData;

const TaskPane: React.FC = () => {
  const [messages, setMessages] = useState<Message[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [selectedData, setSelectedData] = useState<RangeData | null>(null);
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

  useEffect(() => {
    // Initialize Office.js
    const initOffice = async () => {
      try {
        await Office.onReady();
      } catch (error) {
        console.error('Office.js の初期化に失敗しました。ブラウザを再読み込みしてください。', error);
        addMessage('Office.js の初期化に失敗しました。ブラウザを再読み込みしてください。', 'ai', true);
      }
    };

    void initOffice();
  }, []);

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

  const getSelectedData = async (): Promise<RangeData> => {
    try {
      return await Excel.run(async (context: Excel.RequestContext) => {
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
      throw new Error('セル範囲の取得に失敗しました。');
    }
  };

  const handleSendMessage = async (userMessage: string) => {
    addMessage(userMessage, 'user');
    setIsLoading(true);

    try {
      let cellData: RangeData | null = null;
      try {
        cellData = await getSelectedData();
        setSelectedData(cellData);
      } catch (error) {
        console.warn('Selection read failed:', error);
        addMessage('セル範囲が選択されていません。セルを選択してからもう一度お試しください。', 'ai', true);
        setIsLoading(false);
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
        const errorData = await response.json();
        throw new Error(errorData.error || 'API 呼び出しに失敗しました。');
      }

      const result = await response.json();
      addMessage(result.message, 'ai', false, result.action !== 'none');

      if (result.action === 'write' && result.data) {
        try {
          await Excel.run(async (context: Excel.RequestContext) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRange(result.data.address);
            range.values = result.data.values;
            await context.sync();
          });

          addMessage(`${result.data.address} に結果を書き込みました。`, 'ai', false, true);
        } catch (error) {
          console.error('Failed to write to cell:', error);
          addMessage('セルへの書き込みに失敗しました。手動で入力してください。', 'ai', true);
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

  return (
    <div className="chat-container">
      <div className="chat-messages">
        {messages.length === 0 ? (
          <div className="empty-state">
            <div className="empty-state-icon">💬</div>
            <div className="empty-state-title">Excel AI チャットアシスタントチャットアシスタント</div>
            <div className="empty-state-description">
              Excel のセルを選択して、自然言語で指示してください。
              <br />
              データ分析、操作、レポート作成をサポートします。
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

export default TaskPane;





