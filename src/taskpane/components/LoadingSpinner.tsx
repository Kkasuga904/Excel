import React from 'react';

interface LoadingSpinnerProps {
  message?: string;
}

const LoadingSpinner: React.FC<LoadingSpinnerProps> = ({
  message = '処理中...'
}) => {
  return (
    <div className="loading-spinner">
      <div className="spinner"></div>
      <span>{message}</span>
    </div>
  );
};

export default LoadingSpinner;

