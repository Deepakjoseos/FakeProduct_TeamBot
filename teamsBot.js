const { TeamsActivityHandler, TurnContext, CardFactory } = require("botbuilder");

class TeamsBot extends TeamsActivityHandler {
  constructor() {
    super();

    this.onMessage(async (context, next) => {
      console.log("Running with Message Activity.");
      // Check if the message contains a submitted action (product selection)
    if (context.activity.value && context.activity.value.productId) {
      const productId = context.activity.value.productId;
      await this.getProductById(context, productId);
  } else {
      const removedMentionText = TurnContext.removeRecipientMention(context.activity);
      const userInput = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();

      if (!isNaN(userInput) && Number(userInput) > 0) {
        // If input is a number, treat it as product ID
        await this.getProductById(context, userInput);
      } else if (userInput.startsWith("category:")) {
        // If input starts with 'category:', treat it as category search
        const category = userInput.split("category:")[1].trim();
        await this.getProductsByCategory(context, category);
      } else {
        // Otherwise, treat input as a search query
        await this.searchProducts(context, userInput);
      }
    }
      await next();
    });

    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id) {
          await context.sendActivity(
            `Hi there! I'm a Teams bot that displays products based on product ID or category from the API [fakestoreapi.com](https://fakestoreapi.com).\n\n` +
            `You can:\n` +
            `- Type a product ID between 1 and 20 to get details\n` +
            `- Type 'category:electronics' to see products in a category\n` +
            `- Or type a keyword to search for products!`
          );
          break;
        }
      }
      await next();
    });
  }

  // Fetch product by ID
async getProductById(context, productId) {
  try {
    const response = await fetch(`https://fakestoreapi.com/products/${productId}`);
    if (response.ok) {
      const productData = await response.json();

      // Create an Adaptive Card with product details
      const card = {
        type: "AdaptiveCard",
        body: [
          {
            type: "Image",
            url: productData.image,
            altText: "Product Image",
            size: "Medium"
          },
          {
            type: "TextBlock",
            text: productData.title,
            weight: "Bolder",
            size: "Medium"
          },
          {
            type: "TextBlock",
            text: `$${productData.price}`,
            weight: "Bolder",
            size: "Small",
            color: "Good"
          },
          {
            type: "TextBlock",
            text: `**Category**: ${productData.category}`,
            isSubtle: true,
            wrap: true
          },
          {
            type: "TextBlock",
            text: productData.description,
            wrap: true
          },
          {
            type: "TextBlock",
            text: `**Rating**: ${productData.rating.rate} (from ${productData.rating.count} reviews)`,
            wrap: true
          }
        ],
        actions: [],
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
        version: "1.3"
      };

      await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
    } else {
      await context.sendActivity("Invalid product ID. Please try again.");
    }
  } catch (error) {
    console.error(error);
    await context.sendActivity("There was an issue fetching the product. Please try again later.");
  }
}


  // Fetch products by category
async getProductsByCategory(context, category) {
  try {
      const response = await fetch(`https://fakestoreapi.com/products/category/${category}`);
      if (response.ok) {
          const products = await response.json();
          if (products.length > 0) {
              const headerMessage = {
                  type: "TextBlock",
                  text: `Here are the products in the "${category}" category:`,
                  weight: "Bolder",
                  size: "Medium",
                  wrap: true
              };
              // Limit to 2 products for demonstration purposes
              const limitedProducts = products.slice(0, 2);

              const productContainers = limitedProducts.map((product) => ({
                  type: "Container",
                  style: "default", // Simulate border-like effect
                    spacing: "medium",
                    separator: true, // Add a separator to mimic a border
                  items: [
                      {
                          type: "Image",
                          url: product.image,
                          altText: product.title,
                          size: "Medium",
                          horizontalAlignment: "Center"
                      },
                      {
                          type: "TextBlock",
                          text: product.title,
                          weight: "Bolder",
                          wrap: true,
                          horizontalAlignment: "Center"
                      },
                      {
                          type: "TextBlock",
                          text: `$${product.price}`,
                          weight: "Bolder",
                          color: "Good",
                          horizontalAlignment: "Center"
                      },
                      {
                          type: "TextBlock",
                          text: `**Rating:** ${product.rating.rate} (from ${product.rating.count} reviews)`,
                          wrap: true,
                          horizontalAlignment: "Center"
                      },
                      // Add an Action.Submit button to select the product
                      {
                          type: "ActionSet",
                          horizontalAlignment: "Center",
                          actions: [
                              {
                                  type: "Action.Submit",
                                  title: "View Details",
                                  data: { productId: product.id } // Send product ID when clicked
                              }
                          ]
                      }
                  ]
              }));

              const card = {
                  $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
                  type: "AdaptiveCard",
                  version: "1.0",
                  body: [headerMessage, ...productContainers]
              };

              await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
          } else {
              await context.sendActivity(`No products found in the "${category}" category.`);
          }
      } else {
          await context.sendActivity("Invalid category. Please try again.");
      }
  } catch (error) {
      console.error(error);
      await context.sendActivity("There was an issue fetching products by category. Please try again later.");
  }
}

// Search products by keyword
async searchProducts(context, searchTerm) {
  try {
      const response = await fetch(`https://fakestoreapi.com/products`);
      if (response.ok) {
          const products = await response.json();
          const filteredProducts = products.filter((product) =>
              product.title.toLowerCase().includes(searchTerm.toLowerCase())
          );

          if (filteredProducts.length > 0) {
              const headerMessage = {
                  type: "TextBlock",
                  text: `Here are the products that match "${searchTerm}":`,
                  weight: "Bolder",
                  size: "Medium",
                  wrap: true
              };

              const productContainers = filteredProducts.map((product) => ({
                  type: "Container",
                  style: "default", // Simulate border-like effect
                    spacing: "medium",
                    separator: true, // Add a separator to mimic a border
                  items: [
                      {
                          type: "Image",
                          url: product.image,
                          altText: product.title,
                          size: "Medium",
                          horizontalAlignment: "Center"
                      },
                      {
                          type: "TextBlock",
                          text: product.title,
                          weight: "Bolder",
                          wrap: true,
                          horizontalAlignment: "Center"
                      },
                      {
                          type: "TextBlock",
                          text: `$${product.price}`,
                          weight: "Bolder",
                          color: "Good",
                          horizontalAlignment: "Center"
                      },
                      {
                          type: "TextBlock",
                          text: `**Rating:** ${product.rating.rate} (from ${product.rating.count} reviews)`,
                          wrap: true,
                          horizontalAlignment: "Center"
                      },
                      // Add an Action.Submit button to select the product
                      {
                          type: "ActionSet",
                          horizontalAlignment: "Center",
                          actions: [
                              {
                                  type: "Action.Submit",
                                  title: "View Details",
                                  data: { productId: product.id } // Send product ID when clicked
                              }
                          ]
                      }
                  ]
              }));

              const card = {
                  $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
                  type: "AdaptiveCard",
                  version: "1.0",
                  body: [headerMessage, ...productContainers]
              };

              await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
          } else {
              await context.sendActivity(`No products found for "${searchTerm}".`);
          }
      } else {
          await context.sendActivity("There was an issue searching for products. Please try again.");
      }
  } catch (error) {
      console.error(error);
      await context.sendActivity("There was an issue searching for products. Please try again later.");
  }
}

  
  
}

module.exports.TeamsBot = TeamsBot;
