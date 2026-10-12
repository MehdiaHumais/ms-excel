export default function LoadingSkeleton() {
  return (
    <div className="bg-surface rounded-2xl shadow-card p-6 animate-pulse">
      <div className="h-12 w-12 rounded-full bg-gray-200 mb-4 mx-auto"></div>
      <div className="h-6 w-3/4 bg-gray-200 rounded mb-2 mx-auto"></div>
      <div className="h-4 w-1/2 bg-gray-200 rounded mb-2 mx-auto"></div>
      <div className="h-4 w-2/3 bg-gray-200 rounded mx-auto"></div>
    </div>
  );
}