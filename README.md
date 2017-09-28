# jsom-ctx
Simplifies JSOM usage

# Examples
Get web title for current web

```typescript
(async function () {
    try {
        let ctx = new JsomContext(_spPageContextInfo.webAbsoluteUrl);
        ctx = await ctx.load();
        await ExecuteJsomQuery(ctx, [ctx.web, ctx.lists]);
        console.log(ctx.web.get_title());
    } catch (err) {
        console.log(err);
    }
})();
````