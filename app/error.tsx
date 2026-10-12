"use client";

import { useEffect } from "react";

export default function Error({
  error,
  reset
}: {
  error: Error & { digest?: string };
  reset: () => void;
}) {
  // Log the error to an external service if needed
  useEffect(() => {
    console.error(error);
  }, [error]);

  return (
    <div className="flex flex-col items-center justify-center py-20">
      <h2 className="mb-4 text-2xl font-semibold text-primary">
        Something went wrong
      </h2>
      <p className="mb-6 text-center text-sm text-gray-600">
        {error.message}
      </p>
      <button
        onClick={() => reset()}
        className="rounded-2xl bg-primary px-6 py-2 text-sm font-medium text-white transition-colors hover:bg-secondary focus:outline-none focus:ring-2 focus:ring-primary"
      >
        Try again
      </button>
    </div>
  );
}