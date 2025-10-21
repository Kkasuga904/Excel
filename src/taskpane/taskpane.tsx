import React, { useState, useEffect, useRef } from 'react';
 
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

  // 繝｡繝・そ繝ｼ繧ｸ繧定・蜍輔せ繧ｯ繝ｭ繝ｼ繝ｫ
  const scrollToBottom = () => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  };

  useEffect(() => {
    scrollToBottom();
  }, [messages]);

  // Office.js縺ｮ蛻晄悄蛹・  useEffect(() => {
    const initOffice = async () => {
      try {
        await Office.onReady();
        console.log('Office.js initialized');
      } catch (error) {
        console.error('Office.js initialization failed:', error);
        addMessage(
          'Office.js縺ｮ蛻晄悄蛹悶↓螟ｱ謨励＠縺ｾ縺励◆縲ゅヶ繝ｩ繧ｦ繧ｶ繧貞・隱ｭ縺ｿ霎ｼ縺ｿ縺励※縺上□縺輔＞縲・,
          'ai',
          true
        );
      }
    };

    initOffice();
  }, []);

  // 繝｡繝・そ繝ｼ繧ｸ繧定ｿｽ蜉
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

  // 驕ｸ謚槭＆繧後◆繧ｻ繝ｫ繝・・繧ｿ繧貞叙蠕・  const getSelectedData = async () => {
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
      throw new Error('繧ｻ繝ｫ繝・・繧ｿ縺ｮ蜿門ｾ励↓螟ｱ謨励＠縺ｾ縺励◆');
    }
  };

  // 繝｡繝・そ繝ｼ繧ｸ騾∽ｿ｡
  const handleSendMessage = async (userMessage: string) => {
    // 繝ｦ繝ｼ繧ｶ繝ｼ繝｡繝・そ繝ｼ繧ｸ繧定ｿｽ蜉
    addMessage(userMessage, 'user');
    setIsLoading(true);

    try {
      // 驕ｸ謚槭＆繧後◆繧ｻ繝ｫ繝・・繧ｿ繧貞叙蠕・      let cellData = null;
      try {
        cellData = await getSelectedData();
        setSelectedData(cellData);
      } catch (error) {
        console.warn('Could not get selected data:', error);
        addMessage(
          '繝・・繧ｿ縺碁∈謚槭＆繧後※縺・∪縺帙ｓ縲ゅそ繝ｫ繧帝∈謚槭＠縺ｦ縺九ｉ繧ゅ≧荳蠎ｦ縺願ｩｦ縺励￥縺縺輔＞縲・,
          'ai',
          true
        );
        setIsLoading(false);
        return;
      }

      // 繝舌ャ繧ｯ繧ｨ繝ｳ繝峨↓繝ｪ繧ｯ繧ｨ繧ｹ繝磯∽ｿ｡
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
        throw new Error(errorData.error || 'API蜻ｼ縺ｳ蜃ｺ縺励↓螟ｱ謨励＠縺ｾ縺励◆');
      }

      const result = await response.json();

      // AI縺ｮ霑皮ｭ斐ｒ陦ｨ遉ｺ
      addMessage(result.message, 'ai', false, result.action !== 'none');

      // 繧ｻ繝ｫ縺ｫ邨先棡繧呈嶌縺崎ｾｼ縺ｿ
      if (result.action === 'write' && result.data) {
        try {
          await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRange(result.data.address);
            range.values = result.data.values;
            await context.sync();
          });

          addMessage(
            `${result.data.address}縺ｫ邨先棡繧呈嶌縺崎ｾｼ縺ｿ縺ｾ縺励◆`,
            'ai',
            false,
            true
          );
        } catch (error) {
          console.error('Failed to write to cell:', error);
          addMessage(
            '繧ｻ繝ｫ縺ｸ縺ｮ譖ｸ縺崎ｾｼ縺ｿ縺ｫ螟ｱ謨励＠縺ｾ縺励◆縲よ焔蜍輔〒蜈･蜉帙＠縺ｦ縺上□縺輔＞縲・,
            'ai',
            true
          );
        }
      }
    } catch (error) {
      console.error('Error:', error);
      const errorMessage =
        error instanceof Error ? error.message : '繧ｨ繝ｩ繝ｼ縺檎匱逕溘＠縺ｾ縺励◆';
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
            <div className="empty-state-icon">町</div>
            <div className="empty-state-title">Excel AI 繝√Ε繝・ヨ繧｢繧ｷ繧ｹ繧ｿ繝ｳ繝・/div>
            <div className="empty-state-description">
              Excel縺ｮ繧ｻ繝ｫ繧帝∈謚槭＠縺ｦ縲∬・辟ｶ險隱槭〒謖・､ｺ縺励※縺上□縺輔＞縲・              <br />
              繝・・繧ｿ蛻・梵縲∵桃菴懊√Ξ繝昴・繝井ｽ懈・縺ｪ縺ｩ縺悟庄閭ｽ縺ｧ縺吶・            </div>
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
            {isLoading && <LoadingSpinner message="蜃ｦ逅・ｸｭ..." />}
            <div ref={messagesEndRef} />
          </>
        )}
      </div>
      <ChatInput
        onSendMessage={handleSendMessage}
        isLoading={isLoading}
        placeholder="萓・ 縺薙・繝・・繧ｿ繧貞・譫舌＠縺ｦ"
      />
    </div>
  );
};

// React 繧｢繝励Μ繧ｱ繝ｼ繧ｷ繝ｧ繝ｳ繧偵・繧ｦ繝ｳ繝・const root = ReactDOM.createRoot(document.getElementById('root')!);




export default TaskPane;