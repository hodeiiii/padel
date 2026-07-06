import { TournamentPublicPage } from "../../components/TournamentAdmin";

type PublicPageProps = {
  params: Promise<{
    slug: string;
  }>;
};

export default async function PublicPage({ params }: PublicPageProps) {
  const { slug } = await params;

  return <TournamentPublicPage slug={slug} />;
}
