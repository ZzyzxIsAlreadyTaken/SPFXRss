import * as React from "react";
import { useEffect, useState } from "react";
import { Spinner } from "@fluentui/react/lib/Spinner";
import { SPHttpClient } from "@microsoft/sp-http";
import {
  DefaultButton,
  Stack,
  IStackTokens,
  Text,
  Link,
  DocumentCard,
  DocumentCardDetails,
  IDocumentCardStyles,
} from "@fluentui/react";

interface IRSSFeedComponentProps {
  feedUrl: string;
  spHttpClient: SPHttpClient;
  siteUrl: string;
}

interface IChannelImage {
  title: string;
  url: string;
  link: string;
  width?: number;
  height?: number;
}

interface IFeedItem {
  title: string;
  link: string;
  ingress: string;
  image?: string;
  imageTitle?: string;
  imageCredit?: string;
  pubDate?: string;
  creator?: string;
  categories: string[];
}

const stackTokens: IStackTokens = {
  childrenGap: 20,
  padding: 10,
};

const truncateText = (text: string, maxLength: number): string => {
  if (text.length <= maxLength) return text;
  return text.substring(0, maxLength).trim() + "...";
};

const cardStyles: IDocumentCardStyles = {
  root: {
    height: "420px",
    display: "flex",
    flexDirection: "column",
  },
};

const RSSFeedComponent: React.FC<IRSSFeedComponentProps> = ({
  feedUrl,
  spHttpClient,
  siteUrl,
}) => {
  const [items, setItems] = useState<IFeedItem[]>([]);
  const [channelImage, setChannelImage] = useState<IChannelImage | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [currentPage, setCurrentPage] = useState(1);
  const itemsPerPage = 4;

  useEffect(() => {
    const fetchFeed = async () => {
      try {
        const fullFeedUrl = feedUrl.startsWith("http")
          ? feedUrl
          : `https://${feedUrl}`;

        const response = await fetch(fullFeedUrl);
        if (!response.ok) {
          throw new Error(`HTTP error! status: ${response.status}`);
        }
        const text = await response.text();
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(text, "text/xml");

        // Parse channel image
        const imageElement = xmlDoc.querySelector("channel > image");
        if (imageElement) {
          setChannelImage({
            title: imageElement.querySelector("title")?.textContent || "",
            url: imageElement.querySelector("url")?.textContent || "",
            link: imageElement.querySelector("link")?.textContent || "",
            width:
              Number(imageElement.querySelector("width")?.textContent) ||
              undefined,
            height:
              Number(imageElement.querySelector("height")?.textContent) ||
              undefined,
          });
        }

        // Parse items
        const items = Array.from(xmlDoc.querySelectorAll("item")).map(
          (item) => {
            const mediaContent = item.getElementsByTagNameNS(
              "http://search.yahoo.com/mrss/",
              "content"
            )[0];
            const mediaTitle = mediaContent?.getElementsByTagNameNS(
              "http://search.yahoo.com/mrss/",
              "title"
            )[0];
            const mediaCredit = mediaContent?.getElementsByTagNameNS(
              "http://search.yahoo.com/mrss/",
              "credit"
            )[0];

            return {
              title: item.querySelector("title")?.textContent || "",
              link: item.querySelector("link")?.textContent || "",
              ingress: item.querySelector("description")?.textContent || "",
              image:
                mediaContent?.getAttribute("url") ||
                item.querySelector("enclosure")?.getAttribute("url") ||
                undefined,
              imageTitle: mediaTitle?.textContent || undefined,
              imageCredit: mediaCredit?.textContent || undefined,
              pubDate: item.querySelector("pubDate")?.textContent || undefined,
              creator:
                item.querySelector("dc\\:creator")?.textContent || undefined,
              categories: Array.from(item.querySelectorAll("category")).map(
                (cat) => cat.textContent || ""
              ),
            };
          }
        );

        setItems(items);
        setError(null);
      } catch (error) {
        console.error("Error fetching RSS feed:", error);
        setError("Failed to load RSS feed. Please try again later.");
      } finally {
        setLoading(false);
      }
    };

    void fetchFeed();
  }, [feedUrl]);

  if (loading) {
    return <Spinner label="Loading RSS feed..." />;
  }

  if (error) {
    return <div style={{ color: "red" }}>{error}</div>;
  }

  const totalPages = Math.ceil(items.length / itemsPerPage);
  const startIndex = (currentPage - 1) * itemsPerPage;
  const displayedItems = items.slice(startIndex, startIndex + itemsPerPage);

  return (
    <Stack tokens={stackTokens}>
      {channelImage && (
        <Stack.Item align="center">
          <Link href={channelImage.link} target="_blank">
            <img
              src={channelImage.url}
              alt={channelImage.title}
              style={{
                maxWidth: channelImage.width || 114,
                maxHeight: channelImage.height || 114,
              }}
            />
          </Link>
        </Stack.Item>
      )}

      <Stack horizontal wrap tokens={stackTokens}>
        {displayedItems.map((item, index) => (
          <Stack.Item
            key={index}
            grow={1}
            styles={{
              root: {
                minWidth: "280px",
                maxWidth: "600px",
                height: "420px",
                margin: "10px",
              },
            }}
          >
            <DocumentCard styles={cardStyles}>
              {item.image && (
                <div style={{ position: "relative", height: "200px" }}>
                  <img
                    src={item.image}
                    alt={item.imageTitle || item.title}
                    style={{
                      width: "100%",
                      height: "100%",
                      objectFit: "cover",
                    }}
                  />
                  {item.imageCredit && (
                    <div
                      style={{
                        position: "absolute",
                        bottom: 0,
                        right: 0,
                        background: "rgba(0,0,0,0.7)",
                        color: "white",
                        padding: "4px 8px",
                        fontSize: "12px",
                      }}
                    >
                      {item.imageCredit}
                    </div>
                  )}
                </div>
              )}
              <DocumentCardDetails
                styles={{
                  root: {
                    flex: 1,
                    padding: 0,
                  },
                }}
              >
                <Stack
                  tokens={{ childrenGap: 8 }}
                  style={{
                    padding: "12px",
                    height: "100%",
                    display: "flex",
                    flexDirection: "column",
                  }}
                >
                  <Text variant="large" block>
                    <Link href={item.link} target="_blank">
                      {truncateText(item.title, 60)}
                    </Link>
                  </Text>

                  {item.categories.length > 0 && (
                    <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
                      {item.categories.map((category, idx) => (
                        <span
                          key={idx}
                          style={{
                            background: "#f0f0f0",
                            padding: "2px 8px",
                            borderRadius: "12px",
                            fontSize: "12px",
                          }}
                        >
                          {category}
                        </span>
                      ))}
                    </Stack>
                  )}

                  <Text block style={{ flex: 1 }}>
                    {truncateText(item.ingress, 120)}
                  </Text>

                  {(item.creator || item.pubDate) && (
                    <Text
                      variant="small"
                      style={{ color: "#666", marginTop: "auto" }}
                    >
                      {item.creator && <span>By {item.creator}</span>}
                      {item.creator && item.pubDate && <span> • </span>}
                      {item.pubDate && (
                        <span>
                          {new Date(item.pubDate).toLocaleDateString()}
                        </span>
                      )}
                    </Text>
                  )}
                </Stack>
              </DocumentCardDetails>
            </DocumentCard>
          </Stack.Item>
        ))}
      </Stack>

      {totalPages > 1 && (
        <Stack horizontal horizontalAlign="center" tokens={{ childrenGap: 10 }}>
          <DefaultButton
            onClick={() => setCurrentPage((prev) => Math.max(1, prev - 1))}
            disabled={currentPage === 1}
          >
            Previous
          </DefaultButton>
          <Text variant="medium" style={{ lineHeight: "32px" }}>
            Page {currentPage} of {totalPages}
          </Text>
          <DefaultButton
            onClick={() =>
              setCurrentPage((prev) => Math.min(totalPages, prev + 1))
            }
            disabled={currentPage === totalPages}
          >
            Next
          </DefaultButton>
        </Stack>
      )}
    </Stack>
  );
};

export default RSSFeedComponent;
