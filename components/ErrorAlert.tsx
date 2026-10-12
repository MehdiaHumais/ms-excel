type ErrorAlertProps = {
  message: string;
};

export default function ErrorAlert({ message }: ErrorAlertProps) {
  return (
    <div className="bg-red-100 border border-red-400 text-red-700 rounded-2xl p-4 max-w-md mx-auto">
      <p className="font-medium">{message}</p>
    </div>
  );
}