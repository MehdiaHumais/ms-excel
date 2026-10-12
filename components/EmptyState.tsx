export default function EmptyState() {
  return (
    <div className="flex flex-col items-center justify-center h-full text-center p-8">
      <svg
        className="h-16 w-16 text-gray-400 mb-4"
        xmlns="http://www.w3.org/2000/svg"
        fill="none"
        viewBox="0 0 24 24"
        stroke="currentColor"
      >
        <path
          strokeLinecap="round"
          strokeLinejoin="round"
          strokeWidth={2}
          d="M9 12h6M9 16h6M9 8h6"
        />
      </svg>
      <h2 className="text-xl font-semibold mb-2">No workbook opened</h2>
      <p className="text-gray-600">
        Use the File menu to create a new workbook or open an existing one.
      </p>
    </div>
  );
}