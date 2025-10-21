import React from 'react';

interface ChatMessageProps {
  message: string;
  sender: 'user' | 'ai';
  timestamp?: Date;
  isError?: boolean;
  isSuccess?: boolean;
}

const ChatMessage: React.FC<ChatMessageProps> = ({
  message,
  sender,
  timestamp,
  isError = false,
  isSuccess = false
}) => {
  const formatTime = (date: Date) => {
    return date.toLocaleTimeString('ja-JP', {
      hour: '2-digit',
      minute: '2-digit'
    });
  };

  const messageClass = isError ? 'error-message' : isSuccess ? 'success-message' : `message ${sender}`;

  return (
    <div className={messageClass}>
      <div className="message-content">
        {message}
      </div>
      {timestamp && (
        <div className="message-timestamp">
          {formatTime(timestamp)}
        </div>
      )}
    </div>
  );
};

export default ChatMessage;

