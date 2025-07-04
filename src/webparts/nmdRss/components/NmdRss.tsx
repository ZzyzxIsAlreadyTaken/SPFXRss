import * as React from "react";
import { useEffect, useState, useRef } from "react";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
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
import styles from "./NmdRss.module.scss";

interface IRSSFeedComponentProps {
  title: string;
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
    display: "flex",
    flexDirection: "column",
    minWidth: "160px",
    maxWidth: "260px",
  },
};

const RSSFeedComponent: React.FC<IRSSFeedComponentProps> = ({
  feedUrl,
  spHttpClient,
  title,
  siteUrl,
}) => {
  console.log("RSSFeedComponent rendered");
  const [items, setItems] = useState<IFeedItem[]>([]);
  const [channelImage, setChannelImage] = useState<IChannelImage | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [currentPage, setCurrentPage] = useState(1);
  const [imagesLoaded, setImagesLoaded] = useState<Set<number>>(new Set());
  const [itemsPerPage, setItemsPerPage] = useState(6);
  const resizeObserverRef = useRef<ResizeObserver | null>(null);

  // Move updateItemsPerPage to the component scope
  const updateItemsPerPage = React.useCallback(
    (node: HTMLDivElement | null) => {
      if (!node) return;
      const width = node.offsetWidth;
      const cardWidth = 296;
      const gap = 16;
      const cardsPerRow = Math.floor((width + gap) / (cardWidth + gap));
      const perPage = cardsPerRow >= 4 ? 8 : 6;
      setItemsPerPage((prev) => (prev !== perPage ? perPage : prev));
      console.log(
        "Container width:",
        width,
        "Cards per row:",
        cardsPerRow,
        "Items per page:",
        perPage
      );
    },
    []
  );

  // Memoize displayedItems
  const displayedItems = React.useMemo(() => {
    const startIndex = (currentPage - 1) * itemsPerPage;
    return items.slice(startIndex, startIndex + itemsPerPage);
  }, [items, currentPage, itemsPerPage]);

  // Memoize the callback ref
  const mainContainerCallbackRef = React.useCallback(
    (node: HTMLDivElement | null) => {
      console.log("Callback ref called with node:", node);
      if (resizeObserverRef.current) {
        resizeObserverRef.current.disconnect();
        resizeObserverRef.current = null;
      }
      if (node) {
        updateItemsPerPage(node);
        const observer = new window.ResizeObserver(() => {
          console.log("ResizeObserver callback fired");
          updateItemsPerPage(node);
        });
        observer.observe(node);
        resizeObserverRef.current = observer;
      }
    },
    [updateItemsPerPage]
  );

  useEffect(() => {
    console.log(
      "useEffect running. feedUrl:",
      feedUrl,
      "spHttpClient:",
      spHttpClient
    );
    const fetchFeed = async (): Promise<void> => {
      try {
        const fullFeedUrl = feedUrl.startsWith("http")
          ? feedUrl
          : `https://${feedUrl}`;

        console.log("Fetching RSS feed from:", fullFeedUrl);
        console.log("Items per page:", itemsPerPage);

        const response = await spHttpClient.get(
          fullFeedUrl,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: "application/xml, text/xml, */*",
              "Content-Type": "application/xml",
            },
          }
        );

        if (!response.ok) {
          throw new Error(`HTTP error! status: ${response.status}`);
        }
        const text = await response.text();
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(text, "text/xml");

        // Parse channel image
        const imageElement = xmlDoc.querySelector("channel > image");
        if (imageElement) {
          console.log("setChannelImage called");
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
              image: (() => {
                let url =
                  mediaContent?.getAttribute("url") ||
                  item.querySelector("enclosure")?.getAttribute("url") ||
                  undefined;
                if (!url) return undefined;
                // Fix single-slash after protocol (https:/ instead of https://)
                url = url.replace(/^https?:\/(?!\/)/, (match) => match + "/");
                // If the URL starts with http or https, use as is
                if (/^https?:\/\//i.test(url)) return url;
                // If the URL starts with //, prepend https:
                if (/^\/\//.test(url)) return "https:" + url;
                // Otherwise, prepend https://
                return "https://" + url;
              })(),
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

        console.log("setItems called");
        setItems(items);
        console.log("setError called (null)");
        setError(null);
      } catch (error) {
        console.error("Error fetching RSS feed:", error);
        console.log("setError called (error)");
        setError("Failed to load RSS feed. Please try again later.");
      } finally {
        console.log("setLoading called (false)");
        setLoading(false);
      }
    };

    fetchFeed().catch(() => {});
  }, [feedUrl, spHttpClient]);

  // New effect for image preloading with batching
  useEffect(() => {
    let isUnmounted = false;
    let pending = new Set<number>();
    let timeout: number | null = null;

    const update = () => {
      if (isUnmounted) return;
      setImagesLoaded((prev) => {
        const next = new Set(prev);
        for (const idx of Array.from(pending)) {
          next.add(idx);
        }
        return next;
      });
      pending.clear();
      timeout = null;
    };

    // Preload images for the first page only
    const firstPageItems = items.slice(0, itemsPerPage);
    firstPageItems.forEach((item, index) => {
      if (item.image) {
        const img = new window.Image();
        img.onload = () => {
          if (isUnmounted) return;
          pending.add(index);
          if (!timeout) {
            timeout = window.setTimeout(update, 50); // batch updates every 50ms
          }
        };
        img.onerror = () => {
          if (isUnmounted) return;
          pending.add(index);
          if (!timeout) {
            timeout = window.setTimeout(update, 50);
          }
        };
        img.src = item.image;
      }
    });

    // Load remaining images in background
    setTimeout(() => {
      const remainingItems = items.slice(itemsPerPage);
      remainingItems.forEach((item, index) => {
        if (item.image) {
          const img = new window.Image();
          img.onload = () => {
            if (isUnmounted) return;
            const idx = index + itemsPerPage;
            pending.add(idx);
            if (!timeout) {
              timeout = window.setTimeout(update, 50);
            }
          };
          img.onerror = () => {
            if (isUnmounted) return;
            const idx = index + itemsPerPage;
            pending.add(idx);
            if (!timeout) {
              timeout = window.setTimeout(update, 50);
            }
          };
          img.src = item.image;
        }
      });
    }, 1000);

    return () => {
      isUnmounted = true;
      if (timeout) window.clearTimeout(timeout);
    };
  }, [items, itemsPerPage]);

  if (loading) {
    return <Spinner label="Loading RSS feed..." />;
  }

  if (error) {
    return <div style={{ color: "red" }}>{error}</div>;
  }

  const totalPages = Math.ceil(items.length / itemsPerPage);

  return (
    <Stack tokens={stackTokens}>
      <Text variant="large" className={styles.title}>
        {title}
      </Text>
      {channelImage && (
        <Stack.Item align="center">
          <Link href={channelImage.link} target="_blank">
            <img
              src={channelImage.url}
              alt={channelImage.title}
              className={styles.channelImage}
              style={{
                maxWidth: channelImage.width || 114,
                maxHeight: channelImage.height || 114,
              }}
            />
          </Link>
        </Stack.Item>
      )}

      {/* Main container for cards, wrap Stack in a div for ref */}
      <div ref={mainContainerCallbackRef}>
        <Stack
          horizontal
          wrap
          tokens={{ childrenGap: 16, padding: 10 }}
          className={styles.mainContainer}
        >
          {displayedItems.map((item, index) => (
            <Stack.Item key={index} grow className={styles.cardContainer}>
              <DocumentCard styles={cardStyles}>
                {item.image && (
                  <div className={styles.imageContainer}>
                    {imagesLoaded.has(index) ? (
                      <img
                        src={item.image}
                        alt={item.imageTitle || item.title}
                        className={styles.image}
                      />
                    ) : (
                      <div className={styles.imagePlaceholder}>
                        <Spinner size={SpinnerSize.medium} />
                      </div>
                    )}
                    {item.imageCredit && (
                      <div className={styles.imageCredit}>
                        {item.imageCredit}
                      </div>
                    )}
                  </div>
                )}
                <DocumentCardDetails className={styles.cardDetails}>
                  <Stack
                    tokens={{ childrenGap: 8 }}
                    className={styles.cardContent}
                  >
                    <Text
                      variant="large"
                      block
                      className={styles.clampTwoLines}
                    >
                      <Link
                        href={item.link}
                        target="_blank"
                        className={styles.clampTwoLines}
                      >
                        {item.title}
                      </Link>
                    </Text>

                    {item.categories.length > 0 && (
                      <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
                        {item.categories.map((category, idx) => (
                          <span key={idx} className={styles.category}>
                            {category}
                          </span>
                        ))}
                      </Stack>
                    )}

                    <Text block style={{ flex: 1 }}>
                      {truncateText(item.ingress, 120)}
                    </Text>

                    {(item.creator || item.pubDate) && (
                      <Text variant="small" className={styles.metaText}>
                        {item.creator && <span>By {item.creator}</span>}
                        {item.creator && item.pubDate && <span> • </span>}
                        {item.pubDate && (
                          <span className={styles.nmdRssDate}>
                            {(() => {
                              // Parse DD-MM-YYYY to Date object
                              const [day, month, year] =
                                item.pubDate.split("-");
                              const dateObj = new Date(
                                Number(year),
                                Number(month) - 1,
                                Number(day)
                              );
                              // Format in user's local regional format
                              return dateObj.toLocaleDateString(undefined, {
                                year: "numeric",
                                month: "long",
                                day: "numeric",
                              });
                            })()}
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
      </div>

      {totalPages > 1 && (
        <Stack horizontal horizontalAlign="center" tokens={{ childrenGap: 8 }}>
          <DefaultButton
            onClick={() => setCurrentPage((prev) => Math.max(1, prev - 1))}
            disabled={currentPage === 1}
          >
            Previous
          </DefaultButton>
          <Text variant="medium" className={styles.paginationContainer}>
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
