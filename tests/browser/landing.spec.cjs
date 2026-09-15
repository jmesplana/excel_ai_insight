const { test, expect } = require('@playwright/test');

test('landing stays readable across screen sizes and both entry points open the workflow', async ({ page }) => {
    await page.goto('/');
    await expect(page.locator('#landing-page h1')).toHaveText('Your field data.A clearer picture.');
    await expect(page.locator('#step-1')).toBeHidden();
    for (const width of [1440, 390]) {
        await page.setViewportSize({ width, height: 1000 });
        await expect(page.locator('#get-started-btn')).toBeVisible();
        expect(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth)).toBe(true);
        await page.screenshot({ path: `/tmp/aidstack-landing-${width}.png`, fullPage: true });
    }
    await page.locator('#get-started-btn').click();
    await expect(page.locator('#landing-page')).toBeHidden();
    await expect(page.locator('#step-1')).toBeVisible();
    await page.setViewportSize({ width: 1440, height: 1000 });
    await page.locator('#home-link').click();
    await expect(page.locator('#landing-page')).toBeVisible();
    await page.locator('#theme-toggle').click();
    await expect(page.locator('body')).toHaveClass(/dark-mode/);
    await expect(page.locator('body')).toHaveCSS('background-color', 'rgb(15, 23, 42)');
    await page.screenshot({ path: '/tmp/aidstack-landing-dark.png', fullPage: true });
    await page.locator('#start-analyzing-btn').click();
    await expect(page.locator('#step-1')).toBeVisible();
    await expect(page.locator('#landing-page')).toBeHidden();
});
