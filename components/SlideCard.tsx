import Image from "next/image";

export type SlideCardProps = {
  title: string;
  thumbnailUrl: string;
};

export default function SlideCard({ title, thumbnailUrl }: SlideCardProps) {
  return (
    <div className="rounded-2xl bg-white shadow-card transition-transform hover:scale-105">
      <div className="relative h-48 w-full overflow-hidden rounded-t-2xl">
        <Image
          src={thumbnailUrl}
          alt={title}
          fill
          className="object-cover"
          priority={false}
        />
      </div>
      <h3 className="p-4 text-base font-medium text-gray-800">{title}</h3>
    </div>
  );
}