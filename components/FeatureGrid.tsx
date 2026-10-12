import type { FC } from 'react'

export type Feature = {
  title: string
  description: string
  icon: string
}

type Props = {
  features: Feature[]
}

const FeatureCard: FC<{ feature: Feature }> = ({ feature }) => (
  <div className="bg-surface rounded-2xl shadow-card p-6 hover:shadow-glow transition-shadow">
    <div className="text-4xl mb-4">{feature.icon}</div>
    <h3 className="text-2xl font-semibold text-primary mb-2">{feature.title}</h3>
    <p className="text-muted">{feature.description}</p>
  </div>
)

export default function FeatureGrid({ features }: Props) {
  return (
    <section id="features" className="space-y-12">
      <h2 className="text-4xl font-bold text-center text-primary">
        Features
      </h2>
      <div className="grid gap-8 grid-cols-1 md:grid-cols-2 lg:grid-cols-4">
        {features.map((f, i) => (
          <FeatureCard key={i} feature={f} />
        ))}
      </div>
    </section>
  )
}